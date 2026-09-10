"""Regression contracts for pipeline validation and benchmark hardening."""

from __future__ import annotations

import importlib.util
import json
import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch


MODULE_PATH = Path(__file__).with_name("word_revision_pipeline.py")
SPEC = importlib.util.spec_from_file_location("wordrev_pipeline_hardening_test", MODULE_PATH)
assert SPEC and SPEC.loader
PIPELINE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(PIPELINE)


def valid_profile(profile_id: str = "custom-profile") -> dict:
    return {
        "profile_id": profile_id,
        "profile_version": "1.0.0",
        "authority": "test",
        "source_files": [],
        "document_props": {
            "pageWidth": "21cm",
            "pageHeight": "29.7cm",
            "marginTop": "2cm",
            "marginBottom": "2cm",
            "marginLeft": "2cm",
            "marginRight": "2cm",
            "marginHeader": "1cm",
            "marginFooter": "1cm",
        },
        "style_roles": {"body": {}},
        "header": {},
        "footer": {},
        "page_numbering": {},
        "table": {},
        "rules": [{"id": "TEST-001"}],
    }


class PipelineValidationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.source = self.root / "source.docx"
        self.source.write_bytes(b"test")
        self.adapter = self.root / "adapter.py"
        self.adapter.write_text("def configure(job, run_dir):\n    return {}\n", encoding="utf-8")

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def make_job(self, label: str, *, profile: dict | None = None) -> tuple[dict, Path, Path, Path]:
        profile_path = self.root / f"{label}-profile.json"
        profile_payload = profile or valid_profile(f"{label}-profile")
        profile_path.write_text(json.dumps(profile_payload), encoding="utf-8")
        run_dir = self.root / f"{label}-run"
        cache_dir = self.root / f"{label}-cache"
        job = {
            "schema_version": "2.0",
            "job_id": label,
            "scenario": "thesis_format",
            "scope": "chapter",
            "source": str(self.source),
            "adapter": str(self.adapter),
            "profile": str(profile_path),
            "run_dir": str(run_dir),
            "cache_dir": str(cache_dir),
            "notebooklm": "disabled",
            "statistics": {"allow_execution": False, "consume_manifest_only": True},
            "structure_policy": "validate_only",
            "field_policy": "preserve",
            "table_policy": "report_only",
            "style_map": {"body": "/styles/Normal"},
            "output": {"tracked_formatting": False, "delivery_date": "20260905"},
        }
        job_path = self.root / f"{label}.json"
        return job, job_path, run_dir, cache_dir

    def assert_rejected_without_directories(self, label: str, mutate, *, profile=None) -> str:
        job, job_path, run_dir, cache_dir = self.make_job(label, profile=profile)
        mutate(job)
        job_path.write_text(json.dumps(job), encoding="utf-8")
        with self.assertRaises((ValueError, FileNotFoundError)) as caught:
            PIPELINE.Pipeline(job_path)
        self.assertFalse(run_dir.exists(), "validation must precede run_dir creation")
        self.assertFalse(cache_dir.exists(), "validation must precede cache_dir creation")
        return str(caught.exception)

    def test_statistics_contract_is_exact_and_fail_fast(self) -> None:
        for index, value in enumerate((False, None)):
            with self.subTest(consume_manifest_only=value):
                message = self.assert_rejected_without_directories(
                    f"statistics-{index}",
                    lambda job, value=value: job["statistics"].update(
                        {"consume_manifest_only": value}
                    ),
                )
                self.assertIn("consume_manifest_only", message)

    def test_policy_enums_are_fail_fast(self) -> None:
        for index, field in enumerate(("structure_policy", "field_policy", "table_policy")):
            with self.subTest(field=field):
                message = self.assert_rejected_without_directories(
                    f"policy-{index}", lambda job, field=field: job.__setitem__(field, "guess")
                )
                self.assertIn(field, message)

    def test_output_types_and_delivery_date_are_fail_fast(self) -> None:
        cases = (
            ("tracked", lambda job: job["output"].update({"tracked_formatting": "false"}), "tracked_formatting"),
            ("date", lambda job: job["output"].update({"delivery_date": "2026-09-05"}), "delivery_date"),
            ("calendar", lambda job: job["output"].update({"delivery_date": "20260230"}), "delivery_date"),
            ("object", lambda job: job.__setitem__("output", []), "output"),
        )
        for label, mutate, expected in cases:
            with self.subTest(label=label):
                self.assertIn(expected, self.assert_rejected_without_directories(label, mutate))

    def test_style_map_requires_string_selectors(self) -> None:
        message = self.assert_rejected_without_directories(
            "style-map", lambda job: job["style_map"].update({"body": 17})
        )
        self.assertIn("style_map", message)

    def test_valid_minimal_thesis_job_still_initializes(self) -> None:
        job, job_path, run_dir, cache_dir = self.make_job("valid")
        job_path.write_text(json.dumps(job), encoding="utf-8")
        pipeline = PIPELINE.Pipeline(job_path)
        self.assertEqual(pipeline.scenario, "thesis_format")
        self.assertTrue(run_dir.is_dir())
        self.assertTrue(cache_dir.is_dir())

    def test_profile_identity_version_and_required_sections_are_fail_fast(self) -> None:
        malformed_profiles = []
        identity = valid_profile("wrong-id")
        malformed_profiles.append(("identity", identity, "identity mismatch"))
        version = valid_profile("version-profile")
        version["profile_version"] = "latest"
        malformed_profiles.append(("version", version, "profile_version"))
        sections = valid_profile("sections-profile")
        sections.pop("page_numbering")
        malformed_profiles.append(("sections", sections, "required sections"))
        props = valid_profile("props-profile")
        props["document_props"].pop("pageWidth")
        malformed_profiles.append(("props", props, "document_props"))
        for label, profile, expected in malformed_profiles:
            with self.subTest(label=label):
                message = self.assert_rejected_without_directories(
                    label, lambda job: None, profile=profile
                )
                self.assertIn(expected, message)

    def test_profile_geometry_units_are_validated_before_directories(self) -> None:
        invalid = valid_profile("invalid-units-profile")
        invalid["document_props"]["pageWidth"] = "twenty-one cm"
        message = self.assert_rejected_without_directories(
            "invalid-units", lambda job: None, profile=invalid
        )
        self.assertIn("pageWidth", message)
        self.assertIn("supported length", message)

        equivalent = valid_profile("equivalent-units-profile")
        equivalent["document_props"].update({
            "pageWidth": "210mm",
            "pageHeight": "841.8898pt",
            "marginLeft": "1in",
        })
        job, job_path, run_dir, cache_dir = self.make_job(
            "equivalent-units", profile=equivalent
        )
        job_path.write_text(json.dumps(job), encoding="utf-8")
        pipeline = PIPELINE.Pipeline(job_path)
        self.assertEqual(pipeline.scenario, "thesis_format")
        self.assertTrue(run_dir.is_dir())
        self.assertTrue(cache_dir.is_dir())

    def test_profile_geometry_rejects_impossible_margins_before_directories(self) -> None:
        invalid = valid_profile("invalid-margins-profile")
        invalid["document_props"].update({"marginLeft": "11cm", "marginRight": "11cm"})
        message = self.assert_rejected_without_directories(
            "invalid-margins", lambda job: None, profile=invalid
        )
        self.assertIn("horizontal margins", message)

    def test_batch_expansion_rejects_invalid_parent_before_run_dir(self) -> None:
        job, job_path, run_dir, cache_dir = self.make_job("batch-invalid")
        job["scope"] = "chapter_batch"
        job["sources"] = [str(self.source)]
        job["statistics"]["consume_manifest_only"] = False
        job_path.write_text(json.dumps(job), encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "consume_manifest_only"):
            PIPELINE._chapter_jobs(job_path)
        self.assertFalse(run_dir.exists())
        self.assertFalse(cache_dir.exists())

    def test_benchmark_rejects_invalid_job_before_archive_creation(self) -> None:
        job, job_path, run_dir, cache_dir = self.make_job("benchmark-invalid")
        archive = self.root / "benchmark-invalid-archive"
        job["benchmark_dir"] = str(archive)
        job["statistics"]["consume_manifest_only"] = False
        job_path.write_text(json.dumps(job), encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "consume_manifest_only"):
            PIPELINE.benchmark(job_path, 1, 1)
        self.assertFalse(run_dir.exists())
        self.assertFalse(cache_dir.exists())
        self.assertFalse(archive.exists())


class BenchmarkContractTests(unittest.TestCase):
    def test_sample_counts_must_be_positive_before_job_read(self) -> None:
        missing = Path("/definitely/missing/wordrev-job.json")
        for cold, warm in ((0, 1), (-1, 1), (1, 0), (1, -1), (True, 1)):
            with self.subTest(cold=cold, warm=warm):
                with self.assertRaisesRegex(ValueError, "positive integer"):
                    PIPELINE.benchmark(missing, cold, warm)

    def test_directory_allocation_is_collision_proof(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            with patch.object(PIPELINE, "datetime") as frozen_datetime:
                frozen_datetime.now.return_value = datetime(2026, 9, 5, 4, 0, 0)
                first = PIPELINE._benchmark_paths(root / "archive", root / "work")
                second = PIPELINE._benchmark_paths(root / "archive", root / "work")
            self.assertNotEqual(first[0], second[0])
            self.assertNotEqual(first[1], second[1])
            self.assertTrue(first[0].name.startswith("20260905T040000_"))
            self.assertTrue(second[0].name.startswith("20260905T040000_"))
            self.assertTrue(all(path.is_dir() for path in (*first, *second)))

    def test_schema_requires_manifest_only_statistics(self) -> None:
        schema_path = MODULE_PATH.parent.parent / "references" / "job.schema.json"
        schema = json.loads(schema_path.read_text(encoding="utf-8"))
        required = schema["$defs"]["statistics"]["required"]
        self.assertIn("consume_manifest_only", required)


class ManifestAndVisualEvidenceTests(unittest.TestCase):
    def test_runner_timeout_is_recorded_and_reported_as_a_hard_failure(self) -> None:
        runner = PIPELINE.Runner()
        with self.assertRaisesRegex(RuntimeError, "timed out"):
            runner.run(
                [sys.executable, "-c", "import time; time.sleep(10)"],
                timeout=0.1,
            )
        self.assertEqual(len(runner.calls), 1)
        self.assertEqual(runner.calls[0]["returncode"], 124)
        self.assertTrue(runner.calls[0]["timed_out"])

    def test_phase_command_indices_remain_global_across_invocations(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "source.docx"
            source.write_bytes(b"test")
            adapter = root / "adapter.py"
            adapter.write_text("def configure(job, run_dir):\n    return {}\n", encoding="utf-8")
            profile = valid_profile("manifest-profile")
            profile_path = root / "manifest-profile.json"
            profile_path.write_text(json.dumps(profile), encoding="utf-8")
            run_dir = root / "run"
            run_dir.mkdir()
            manifest = {
                "schema_version": "1.0",
                "job_id": "manifest",
                "job_path": str(root / "job.json"),
                "created_at": "2026-09-05T00:00:00Z",
                "updated_at": "2026-09-05T00:00:00Z",
                "notebooklm": "disabled",
                "scenario": "thesis_format",
                "phases": {},
                "commands": [{"command": ["old"], "duration_seconds": 1, "returncode": 0}],
            }
            (run_dir / "run_manifest.json").write_text(json.dumps(manifest), encoding="utf-8")
            job = {
                "schema_version": "2.0", "job_id": "manifest", "scenario": "thesis_format",
                "scope": "chapter", "source": str(source), "adapter": str(adapter),
                "profile": str(profile_path), "run_dir": str(run_dir),
                "cache_dir": str(root / "cache"), "notebooklm": "disabled",
                "statistics": {"allow_execution": False, "consume_manifest_only": True},
                "structure_policy": "validate_only", "field_policy": "preserve",
                "table_policy": "report_only", "style_map": {"body": "/styles/Normal"},
                "output": {"tracked_formatting": False, "delivery_date": "20260905"},
            }
            job_path = root / "job.json"
            job_path.write_text(json.dumps(job), encoding="utf-8")
            pipeline = PIPELINE.Pipeline(job_path)

            def action():
                pipeline.runner.calls.append({
                    "command": ["new"], "duration_seconds": 0.1, "returncode": 0
                })
                return [source]

            record = pipeline._phase("probe", "fingerprint", action, resume=False)
            self.assertEqual(record["command_start_index"], 1)
            self.assertEqual(record["command_end_index"], 2)
            saved = json.loads((run_dir / "run_manifest.json").read_text(encoding="utf-8"))
            self.assertEqual(len(saved["commands"]), 2)

    def test_manual_review_pages_include_body_and_orientation_boundaries(self) -> None:
        page_sizes = [
            {"page": page, "orientation": "landscape" if 35 <= page <= 38 else "portrait"}
            for page in range(1, 40)
        ]
        pages = PIPELINE._manual_review_pages(
            39, [10], {"page_sizes": page_sizes}
        )
        self.assertEqual(pages, [1, 2, 3, 10, 34, 35, 38, 39])


if __name__ == "__main__":
    unittest.main(verbosity=2)
