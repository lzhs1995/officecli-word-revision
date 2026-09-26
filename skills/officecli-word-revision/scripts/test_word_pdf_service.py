"""Word PDF service contracts that do not require a live Word session."""

from __future__ import annotations

import importlib.util
import tempfile
import subprocess
import unittest
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

    def _run_fixture(self, failure=None):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        root = Path(temporary.name)
        source = root / 'source.docx'
        source.write_bytes(b'original unchanged')
        output = root / 'output.pdf'
        stages = []
        with patch('pathlib.Path.home', return_value=root), patch('shutil.which', return_value='/bin/tool'):
            def runner(command, **kwargs):
                if command[0].endswith('osascript'):
                    stage = Path(command[1]).stem.upper()
                    stages.append(stage)
                    if stage == failure:
                        return subprocess.CompletedProcess(command, 1, '', 'AppleEvent timeout (-1712)')
                    if stage == 'EXPORT':
                        session.temporary.write_bytes(b'%PDF-test')
                    return subprocess.CompletedProcess(command, 0, 'READY', '')
                if command[0] == 'pdfinfo':
                    return subprocess.CompletedProcess(command, 0, 'Pages: 2\nPage size: 595.28 x 841.89 pts\n', '')
                return subprocess.CompletedProcess(command, 0, 'First page\fLast page', '')
            session = module.ExportSession(source, output, runner=runner)
            result = session.export()
        return root, session, stages, result

    def test_success_uses_container_and_publishes_before_owned_close(self):
        root, session, stages, result = self._run_fixture()
        self.assertTrue(result['success'])
        self.assertEqual(result['cleanup'], 'PASS')
        self.assertIn('Containers/com.microsoft.Word/Data/Documents', str(session.owned))
        self.assertEqual(session.output.read_bytes(), b'%PDF-test')
        self.assertFalse(session.staging.exists())
        self.assertEqual(stages[-1], 'CLOSE')
        self.assertTrue(result['input_unchanged'])

    def test_export_timeout_never_closes_retries_or_publishes(self):
        root, session, stages, result = self._run_fixture('EXPORT')
        self.assertFalse(result['success'])
        self.assertEqual(result['failed_stage'], 'EXPORT')
        self.assertNotIn('CLOSE', stages)
        self.assertEqual(stages.count('EXPORT'), 1)
        self.assertFalse(session.output.exists())
        self.assertTrue((root / '.cache/wordrev/word-automation.pending.json').exists())

    def test_cleanup_timeout_keeps_verified_pdf_and_pending_state(self):
        root, session, stages, result = self._run_fixture('CLOSE')
        self.assertTrue(result['success'])
        self.assertEqual(result['code'], 'EXPORT_PASS_CLEANUP_PENDING')
        self.assertEqual(result['cleanup'], 'PENDING')
        self.assertTrue(session.output.exists())
        self.assertTrue((root / '.cache/wordrev/word-automation.pending.json').exists())


if __name__ == "__main__":
    unittest.main()
