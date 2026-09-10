from __future__ import annotations

import argparse
import importlib.util
import json
import tempfile
import unittest
from pathlib import Path


HERE = Path(__file__).resolve().parent
PIPELINE_PATH = HERE / "word_revision_pipeline.py"
SPEC = importlib.util.spec_from_file_location("wordrev_pipeline", PIPELINE_PATH)
if SPEC is None or SPEC.loader is None:
    raise ImportError(PIPELINE_PATH)
PIPELINE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(PIPELINE)


class WordRevisionContractTests(unittest.TestCase):
    def _source(self, root: Path, name: str = "source.docx") -> Path:
        path = root / name
        path.write_bytes(b"test fixture")
        return path

    def _adapter(self, root: Path) -> Path:
        path = root / "adapter.py"
        path.write_text(
            "def configure(job, run_dir):\n"
            "    return {'prepared': str(run_dir / 'prepared.docx')}\n",
            encoding="utf-8",
        )
        return path

    def test_schema_1_job_routes_to_manuscript_without_rewrite(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = self._source(root)
            content = self._source(root, "content.docx")
            adapter = self._adapter(root)
            job = {
                "schema_version": "1.0",
                "job_id": "legacy",
                "source": str(source),
                "content_source": str(content),
                "adapter": str(adapter),
                "run_dir": str(root / "run"),
                "cache_dir": str(root / "cache"),
                "notebooklm": "disabled",
                "statistics": {"allow_execution": False, "consume_manifest_only": True},
            }
            job_path = root / "job.json"
            job_path.write_text(json.dumps(job), encoding="utf-8")
            pipeline = PIPELINE.Pipeline(job_path)
            self.assertEqual(pipeline.scenario, "manuscript_revision")
            self.assertNotIn("scenario", json.loads(job_path.read_text(encoding="utf-8")))

    def test_thesis_init_writes_schema_2_contract(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = self._source(root)
            output = root / "job.json"
            args = argparse.Namespace(
                source=[str(source)],
                scenario="thesis-format",
                scope="chapter",
                run_dir=str(root / "run"),
                cache_dir=str(root / "cache"),
                word_pdf_staging_dir=str(root / "word"),
                job_id=None,
                style=["body=/styles/Normal"],
                profile="ruc-doctoral-2026",
                thesis_title=None,
                content_source=None,
                adapter=None,
                revision_author="Codex",
                out=str(output),
            )
            PIPELINE.init_job(args)
            job = json.loads(output.read_text(encoding="utf-8"))
            self.assertEqual(job["schema_version"], "2.0")
            self.assertEqual(job["scenario"], "thesis_format")
            self.assertEqual(job["scope"], "chapter")
            self.assertFalse(job["output"]["tracked_formatting"])
            self.assertRegex(job["output"]["delivery_date"], r"^[0-9]{8}$")
            self.assertIn("benchmark_dir", job)
            self.assertEqual(job["notebooklm"], "disabled")
            self.assertFalse(job["statistics"]["allow_execution"])

    def test_manuscript_init_opts_into_generic_journal_layout(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = self._source(root)
            content = self._source(root, "content.docx")
            adapter = self._adapter(root)
            output = root / "job.json"
            args = argparse.Namespace(
                source=[str(source)],
                scenario="manuscript-revision",
                scope=None,
                run_dir=str(root / "run"),
                cache_dir=str(root / "cache"),
                word_pdf_staging_dir=str(root / "word"),
                job_id=None,
                style=[],
                profile=None,
                thesis_title=None,
                content_source=str(content),
                adapter=str(adapter),
                revision_author="Codex",
                out=str(output),
            )
            PIPELINE.init_job(args)
            job = json.loads(output.read_text(encoding="utf-8"))
            self.assertEqual(job["scenario"], "manuscript_revision")
            self.assertEqual(job["publication_layout"]["mode"], "generic_journal")
            self.assertEqual(
                job["publication_layout"]["table"]["alignment_mode"],
                "han-left-short-nonhan-right",
            )
            self.assertTrue(
                job["publication_layout"]["pagination"]["require_narrative_on_object_pages"]
            )

    def test_chapter_batch_expands_without_merge(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            sources = [self._source(root, f"chapter{index}.docx") for index in (1, 2)]
            job = {
                "schema_version": "2.0",
                "job_id": "batch",
                "scenario": "thesis_format",
                "scope": "chapter_batch",
                "source": str(sources[0]),
                "sources": [str(path) for path in sources],
                "profile": "ruc-doctoral-2026",
                "adapter": str(HERE / "thesis_format_adapter.py"),
                "run_dir": str(root / "run"),
                "cache_dir": str(root / "cache"),
                "notebooklm": "disabled",
                "statistics": {"allow_execution": False, "consume_manifest_only": True},
                "style_map": {"body": "/styles/Normal"},
                "output": {"tracked_formatting": False},
            }
            job_path = root / "job.json"
            job_path.write_text(json.dumps(job), encoding="utf-8")
            children = PIPELINE._chapter_jobs(job_path)
            self.assertEqual(len(children), 2)
            for child_path, _ in children:
                child = json.loads(child_path.read_text(encoding="utf-8"))
                self.assertEqual(child["scope"], "chapter")
                self.assertNotIn("sources", child)
                self.assertNotIn("merge", child)

    def test_execution_guardrails_reject_forbidden_branches(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = self._source(root)
            adapter = self._adapter(root)
            base = {
                "schema_version": "2.0",
                "job_id": "guardrail",
                "scenario": "thesis_format",
                "scope": "chapter",
                "source": str(source),
                "adapter": str(adapter),
                "profile": "ruc-doctoral-2026",
                "style_map": {"body": "/styles/Normal"},
                "output": {"tracked_formatting": False},
                "run_dir": str(root / "run"),
                "cache_dir": str(root / "cache"),
                "notebooklm": "disabled",
                "statistics": {"allow_execution": False, "consume_manifest_only": True},
            }
            for key, value in (("notebooklm", "enabled"), ("statistics", {"allow_execution": True})):
                job = dict(base)
                job[key] = value
                job_path = root / f"{key}.json"
                job_path.write_text(json.dumps(job), encoding="utf-8")
                with self.assertRaises(ValueError):
                    PIPELINE.Pipeline(job_path)

    def test_ruc_profile_is_versioned_and_uses_targeted_refresh(self):
        profile_path = HERE.parent / "references" / "profiles" / "ruc-doctoral-2026.json"
        profile = json.loads(profile_path.read_text(encoding="utf-8"))
        self.assertEqual(profile["profile_id"], "ruc-doctoral-2026")
        self.assertRegex(profile["profile_version"], r"^\d+\.\d+\.\d+$")
        self.assertNotIn("updateFields", profile["document_props"])
        self.assertNotIn("recalcFields", profile["document_props"])
        self.assertTrue(profile["document_props"]["mirrorMargins"])

    def test_thesis_layout_sidecars_restore_from_metadata(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            spec = importlib.util.spec_from_file_location(
                "wordrev_thesis_adapter", HERE / "thesis_format_adapter.py"
            )
            self.assertIsNotNone(spec)
            self.assertIsNotNone(spec.loader)
            adapter = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(adapter)
            paths = {
                "semantic_before": str(root / "evidence" / "semantic_before.json"),
                "command_count": str(root / "evidence" / "command_count.txt"),
            }
            metadata = {
                "semantic_before": {"body_text_sha256": "abc"},
                "officecli_commands": 7,
            }
            adapter.restore_layout_sidecars(metadata, paths)
            restored = json.loads(
                Path(paths["semantic_before"]).read_text(encoding="utf-8")
            )
            self.assertEqual(restored["body_text_sha256"], "abc")
            self.assertEqual(
                Path(paths["command_count"]).read_text(encoding="utf-8"), "7"
            )


if __name__ == "__main__":
    unittest.main()
