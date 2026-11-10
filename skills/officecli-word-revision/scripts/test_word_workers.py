"""Worker aliases must not bypass legacy instance locks or broaden output ownership."""
import json
import tempfile
import threading
import unittest
from pathlib import Path
from unittest.mock import patch

from word_workers import resolve_worker, check_output
from word_runtime import word_lock, mark_word_pending
from word_windows_native import native_script


class WorkerTests(unittest.TestCase):
    def test_registered_alias_uses_legacy_lock_and_word_only_instance(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            identity = dict(machine='test-mac', user='fixture', session='gui/501', platform='Darwin')
            row = dict(identity, word_instance='Microsoft Word/current-user-session',
                       zotero_instance=None, output_root=str(root))
            registry = root / 'workers.json'
            registry.write_text(json.dumps(dict(schema=1, workers={'mac-test': row})))
            with patch('word_workers.local_identity', return_value=identity), patch('pathlib.Path.home', return_value=root):
                registered = resolve_worker('mac-test', registry)
                self.assertEqual(registered['native_lock'], resolve_worker()['native_lock'])
                check_output(registered, root / 'owned.pdf')
                with self.assertRaises(ValueError):
                    check_output(registered, root.parent / 'foreign.pdf')
                for field, value in [('native_lock', str(root / 'bypass.lock')), ('user', 'foreign'),
                                     ('zotero_instance', 'untested-profile')]:
                    registry.write_text(json.dumps(dict(schema=1, workers={'mac-test': dict(row, **{field:value})})))
                    with self.assertRaises(ValueError):
                        resolve_worker('mac-test', registry)

    def test_word_only_and_legacy_word_zotero_conflict(self):
        with tempfile.TemporaryDirectory() as tmp, patch('pathlib.Path.home', return_value=Path(tmp)):
            errors = []
            def competitor():
                try:
                    with word_lock('word-only', timeout=.05):
                        errors.append('unexpected entry')
                except TimeoutError:
                    errors.append('conflict')
            with word_lock('legacy-word-zotero'):
                t = threading.Thread(target=competitor); t.start(); t.join()
            self.assertEqual(errors, ['conflict'])
            mark_word_pending({'status':'UNKNOWN', 'owner':'legacy'})
            with self.assertRaisesRegex(RuntimeError, 'WORD_RECOVERY_REQUIRED'):
                with word_lock('word-only'):
                    self.fail('entered pending native instance')

    def test_windows_script_preserves_foreign_processes_and_has_exact_paths(self):
        script = native_script(r'C:\owned\中文.docx', r'C:\owned\result.pdf', r'C:\owned\receipt.json',
                               dict(machine='TEST-WINDOWS', user='fixture-user', session='7'))
        self.assertIn("Documents.Open('C:\\owned\\中文.docx'", script)
        self.assertIn('OWNED_PATH_MISMATCH', script)
        self.assertIn('WORD_INSTANCE_BUSY', script)
        self.assertIn('GetWindowThreadProcessId', script)
        self.assertNotIn('Stop-Process', script)
        self.assertNotIn('@OWNED@', script)
        self.assertNotIn('@USER@', script)


if __name__ == '__main__':
    unittest.main()
