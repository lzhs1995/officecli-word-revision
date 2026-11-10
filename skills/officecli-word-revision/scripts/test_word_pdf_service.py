"""Word PDF service contracts that do not require a live Word session."""

from __future__ import annotations

import importlib.util
import tempfile
import subprocess
import unittest
import json
import threading
import os
from pathlib import Path
from unittest.mock import patch


ROOT = Path(__file__).parent
spec = importlib.util.spec_from_file_location("word_pdf_service", ROOT / "word_pdf_service.py")
module = importlib.util.module_from_spec(spec)
assert spec.loader is not None
spec.loader.exec_module(module)


class WordPdfServiceTest(unittest.TestCase):
    def test_existing_output_is_preserved_before_word_preflight(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "source.docx"
            output = root / "submission.pdf"
            source.write_bytes(b"docx")
            output.write_bytes(b"%PDF-known-good")
            result = module.ExportSession(source, output, runner=lambda *args, **kwargs: None).export()
            self.assertFalse(result["success"])
            self.assertEqual(result["code"], "OUTPUT_EXISTS")
            self.assertEqual(output.read_bytes(), b"%PDF-known-good")

    def test_open_stage_uses_direct_posix_file_and_writes_evidence(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "source.docx"
            output = root / "submission.pdf"
            source.write_bytes(b"docx")
            calls = []

            class Process:
                returncode = 0
                stdout = ""
                stderr = ""

            def runner(command, **kwargs):
                calls.append(command)
                return Process()

            session = module.ExportSession(source, output, runner=runner)
            session.apple("OPEN", "open (POSIX file \"/tmp/source.docx\") read only true add to recent files false")
            script = (session.evidence / "open.applescript").read_text()
            self.assertIn("open (POSIX file", script)
            self.assertNotIn("open file name inputFile", script)
            self.assertTrue(calls)

    def test_binding_requires_owned_full_path(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.docx"
            source.write_bytes(b"docx")
            session = module.ExportSession(source, Path(tmp) / "out.pdf")
            binding = session.binding()
            self.assertIn("posix full name of docRef", binding)
            self.assertIn(str(session.owned), binding)

    def _run_fixture(self, failure=None, qa_failure=False, mutate_source=False,
                     produce_pdf=True, native_only=False):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        root = Path(temporary.name)
        source = root / 'source.docx'
        source.write_bytes(b'original unchanged')
        output = root / 'output.pdf'
        stages = []
        with patch('pathlib.Path.home', return_value=root), patch('shutil.which', return_value='/bin/tool'):
            def runner(command, **kwargs):
                if command[0].endswith('pgrep'):
                    return subprocess.CompletedProcess(command, 0, '9001\n', '')
                if command[0].endswith('/ps'):
                    return subprocess.CompletedProcess(command, 0, str(os.getuid()) + ' Tue fixture Microsoft Word', '')
                if command[0].endswith('osascript'):
                    stage = Path(command[1]).stem.upper()
                    stages.append(stage)
                    if stage == failure:
                        return subprocess.CompletedProcess(command, 1, '', 'AppleEvent timeout (-1712)')
                    if stage == 'EXPORT' and produce_pdf:
                        session.temporary.write_bytes(b'%PDF-test')
                    return subprocess.CompletedProcess(command, 0, 'READY', '')
                if command[0] == 'pdfinfo':
                    # Pure QA runs after native CLOSE and outside the actual OS lock.
                    self.assertEqual(stages[-1], 'CLOSE')
                    entered = []
                    def competing_writer():
                        from word_runtime import word_lock
                        with word_lock('test-qa-boundary', timeout=.2):
                            entered.append(True)
                    thread = threading.Thread(target=competing_writer)
                    thread.start(); thread.join()
                    self.assertEqual(entered, [True])
                    if mutate_source:
                        source.write_bytes(b'changed')
                    if qa_failure:
                        return subprocess.CompletedProcess(command, 1, '', 'invalid PDF')
                    return subprocess.CompletedProcess(command, 0, 'Pages: 2\nPage size: 595.28 x 841.89 pts\n', '')
                return subprocess.CompletedProcess(command, 0, 'First page\fLast page', '')
            session = module.ExportSession(source, output, runner=runner)
            result = session.export(native_only=native_only)
        return root, session, stages, result

    def test_success_closes_before_qa_and_publishes_outside_native_lock(self):
        root, session, stages, result = self._run_fixture()
        self.assertTrue(result['success'])
        self.assertEqual(result['cleanup'], 'PASS')
        self.assertIn('Containers/com.microsoft.Word/Data/Documents', str(session.owned))
        self.assertEqual(session.output.read_bytes(), b'%PDF-test')
        self.assertFalse(session.staging.exists())
        self.assertEqual(stages[-1], 'CLOSE')
        self.assertTrue(result['input_unchanged'])
        self.assertTrue(result['native_closed'])
        self.assertGreaterEqual(result['os_lock_held_seconds'], 0)
        self.assertIsNone(result['logical_reservation_seconds'])
        self.assertFalse(session.claim.exists())

    def test_export_timeout_never_closes_retries_or_publishes(self):
        root, session, stages, result = self._run_fixture('EXPORT')
        self.assertFalse(result['success'])
        self.assertEqual(result['failed_stage'], 'EXPORT')
        self.assertNotIn('CLOSE', stages)
        self.assertEqual(stages.count('EXPORT'), 1)
        self.assertFalse(session.output.exists())
        self.assertTrue((root / '.cache/wordrev/word-automation.pending.json').exists())

    def test_close_timeout_retains_native_output_without_qa_or_publication(self):
        root, session, stages, result = self._run_fixture('CLOSE')
        self.assertFalse(result['success'])
        self.assertEqual(result['code'], 'TIMEOUT_UNRESOLVED')
        self.assertEqual(result['cleanup'], 'PENDING')
        self.assertFalse(session.output.exists())
        self.assertTrue(session.temporary.exists())
        self.assertTrue(session.claim.exists())
        self.assertTrue((root / '.cache/wordrev/word-automation.pending.json').exists())

    def test_qa_failure_is_closed_and_does_not_block_next_native_writer(self):
        root, session, stages, result = self._run_fixture(qa_failure=True)
        self.assertEqual(result['code'], 'NATIVE_CLOSED_QA_FAILED')
        self.assertEqual(result['cleanup'], 'PASS')
        self.assertTrue(result['native_closed'])
        self.assertFalse(session.output.exists())
        self.assertTrue(session.frozen_pdf.exists())
        self.assertFalse(session.claim.exists())
        self.assertFalse((root / '.cache/wordrev/word-automation.pending.json').exists())

    def test_source_changed_after_native_close_rejects_publication_without_pending(self):
        root, session, stages, result = self._run_fixture(mutate_source=True)
        self.assertEqual(result['code'], 'SOURCE_CHANGED')
        self.assertFalse(result['success'])
        self.assertTrue(result['native_closed'])
        self.assertFalse(session.output.exists())
        self.assertFalse((root / '.cache/wordrev/word-automation.pending.json').exists())

    def test_same_output_claim_rejects_before_word(self):
        with tempfile.TemporaryDirectory() as tmp, patch('shutil.which', return_value='/bin/tool'):
            root = Path(tmp); source = root / 's.docx'; source.write_bytes(b'unchanged')
            output = root / 'o.pdf'; claim = root / '.o.pdf.wordrev-claim'
            claim.write_text(json.dumps({'run_id': 'another-owner'}))
            result = module.ExportSession(source, output, runner=lambda *a, **k: self.fail('native called')).export()
            self.assertIn('OUTPUT_CLAIMED', result['error'])
            self.assertEqual(json.loads(claim.read_text())['run_id'], 'another-owner')

    def test_unknown_worker_rejected_before_staging(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / 's.docx'; source.write_bytes(b'unchanged')
            with self.assertRaisesRegex(ValueError, 'UNKNOWN_WORKER'):
                module.ExportSession(source, Path(tmp) / 'o.pdf', worker_id='arbitrary-new-instance')

    def test_native_only_freeze_failure_has_nonzero_cli_exit(self):
        _, session, _, result = self._run_fixture(produce_pdf=False, native_only=True)
        self.assertTrue(result['native_closed'])
        self.assertEqual(result['code'], 'NATIVE_CLOSED_QA_FAILED')
        self.assertFalse(session.frozen_pdf.exists())
        self.assertFalse((session.evidence / 'native-closed.json').exists())
        with patch('sys.argv', ['service', 'input.docx', 'out.pdf', '--native-only']), \
                patch.object(module, 'export_pdf', return_value=result), patch('builtins.print'):
            self.assertEqual(module.main(), 1)

    def test_native_only_frozen_phase_can_exit_zero_without_qa_pass(self):
        _, session, _, result = self._run_fixture(native_only=True)
        self.assertFalse(result['success'])
        self.assertEqual(result['code'], 'NATIVE_CLOSED_AWAITING_QA')
        self.assertTrue(session.frozen_pdf.exists())
        self.assertTrue((session.evidence / 'native-closed.json').exists())
        with patch('sys.argv', ['service', 'input.docx', 'out.pdf', '--native-only']), \
                patch.object(module, 'export_pdf', return_value=result), patch('builtins.print'):
            self.assertEqual(module.main(), 0)

    def test_windows_dispatch_exception_preserves_pending(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / 'source.docx'; source.write_bytes(b'unchanged')
            worker = dict(backend='word-com', machine='TEST-WINDOWS', user='fixture-user', session='7')
            with patch('pathlib.Path.home', return_value=root), \
                    patch.dict(os.environ, {'LOCALAPPDATA': str(root)}), \
                    patch('word_workers.resolve_worker', return_value=worker), \
                    patch('word_windows_native.run_native', side_effect=UnicodeError('synthetic decoding failure')):
                result = module.ExportSession(source, root / 'out.pdf').export(native_only=True)
            self.assertFalse(result['native_closed'])
            self.assertEqual(result['cleanup'], 'PENDING')
            self.assertEqual(result['failed_stage'], 'NATIVE_DISPATCH_UNKNOWN')
            pending = root / '.cache/wordrev/word-automation.pending.json'
            self.assertTrue(pending.exists())
            self.assertTrue(json.loads(pending.read_text())['native_result']['uncertain'])


if __name__ == "__main__":
    unittest.main()
