"""Small fail-closed regression tests; never open Word or run statistical models."""
import hashlib
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock

from word_revision_pipeline import analysis_artifacts, Pipeline
from word_runtime import word_lock, owned_script


class ArtifactTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.file = self.root / 'results.csv'
        self.file.write_text('chain,estimate\nAA-S,1.12\n')
        self.item = {'path': 'results.csv', 'sha256': hashlib.sha256(self.file.read_bytes()).hexdigest()}

    def tearDown(self): self.tmp.cleanup()

    def test_valid_and_legacy_hash_bound(self):
        for contract in [{'artifacts':[self.item]}, {'outputs':{'results':self.item}}]:
            self.assertEqual(analysis_artifacts(contract,self.root)[0]['path'],str(self.file.resolve()))

    def test_unhashed_legacy_is_not_silently_skipped(self):
        with self.assertRaises(ValueError): analysis_artifacts({'tables':{'results':'results.csv'}},self.root)

    def test_missing_hash_and_duplicates_fail(self):
        for items in [[{'path':'results.csv'}],[self.item,self.item]]:
            with self.assertRaises(ValueError): analysis_artifacts({'artifacts':items},self.root)

    def test_changed_or_missing_result_fails(self):
        self.file.write_text('changed')
        with self.assertRaises(RuntimeError): analysis_artifacts({'artifacts':[self.item]},self.root)
        self.file.unlink()
        with self.assertRaises(RuntimeError): analysis_artifacts({'artifacts':[self.item]},self.root)

    def test_locked_owner_cannot_be_overwritten(self):
        path=self.root/'word.lock'
        with word_lock('first',path=path):
            with self.assertRaises(TimeoutError):
                with word_lock('second',timeout=.05,path=path): pass
            self.assertIn('first',path.read_text())

    def test_failed_export_retains_owned_staging_for_recovery(self):
        p=Pipeline.__new__(Pipeline)
        p.job={'job_id':'owned','word_pdf_staging_dir':str(self.root/'staging')}
        p.adapter=Mock()
        p.adapter.export_pdf.side_effect=RuntimeError('Word modal')
        with self.assertRaises(RuntimeError):p._export_pdf(self.file,self.root/'out.pdf','layout')
        self.assertEqual(len(list((self.root/'staging').glob('*.docx'))),1)

    def test_refresh_never_closes_all_documents(self):
        script=owned_script(self.file,'refresh')
        self.assertIn('saved of targetDocument is false',script)
        self.assertIn('dirty state is not completion proof',script)
        self.assertNotIn('close every document',script)
        self.assertNotIn('delay 5',script)
        self.assertNotIn('tell application "NotebookLM"',script)

if __name__=='__main__': unittest.main()
