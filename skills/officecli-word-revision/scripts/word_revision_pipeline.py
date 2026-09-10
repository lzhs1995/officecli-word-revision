#!/usr/bin/env python3
"""Hash-driven Word tracked-revision pipeline built around OfficeCLI and Word."""

from __future__ import annotations

import argparse
import concurrent.futures
import csv
import hashlib
import importlib.util
import json
import os
import signal
import shutil
import statistics
import subprocess
import sys
import tempfile
import time
import uuid
import traceback
import re
import inspect
from contextlib import nullcontext
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable


RTK = shutil.which("rtk") or "/Users/lzhs/.local/bin/rtk"
OFFICECLI = shutil.which("officecli") or "/Users/lzhs/.local/bin/officecli"
SKILL_DIR = Path(__file__).resolve().parent.parent
THESIS_ADAPTER = SKILL_DIR / "scripts" / "thesis_format_adapter.py"
PROFILE_DIR = SKILL_DIR / "references" / "profiles"
SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))
from journal_layout import default_publication_layout_policy, normalize_publication_layout_policy
from word_runtime import word_lock
from native_table_fit import native_fit, transfer_geometry

THESIS_POLICY_VALUES = {
    "structure_policy": {"validate_only", "explicit_anchors"},
    "field_policy": {"preserve", "repair", "create"},
    "table_policy": {"report_only", "explicit", "all_body_tables"},
}
PROFILE_REQUIRED_SECTIONS = {
    "profile_id", "profile_version", "document_props", "style_roles",
    "header", "footer", "page_numbering", "table", "rules",
}
PROFILE_REQUIRED_DOCUMENT_PROPS = {
    "pageWidth", "pageHeight", "marginTop", "marginBottom", "marginLeft",
    "marginRight", "marginHeader", "marginFooter",
}
PROFILE_LENGTH_TO_CM = {
    "cm": 1.0,
    "mm": 0.1,
    "in": 2.54,
    "pt": 2.54 / 72.0,
}


def utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def stable_hash(value: Any) -> str:
    payload = json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def analysis_artifacts(contract: dict, base: Path) -> list[dict]:
    """Validate canonical artifacts or a fully hash-bound legacy 1.0 manifest."""
    if "artifacts" in contract:
        entries = contract["artifacts"]
    else:
        if "tables" in contract or "scripts" in contract:
            raise ValueError("Unhashed legacy analysis manifest: create a versioned artifacts sidecar; original is not rewritten")
        entries = [contract[k] for k in ("data", "script") if k in contract]
        entries += list(contract.get("outputs", {}).values())
    if not isinstance(entries, list) or not entries:
        raise ValueError("Analysis manifest must declare non-empty hash-bound artifacts")
    results, seen = [], set()
    for entry in entries:
        if not isinstance(entry, dict) or not entry.get("path") or not re.fullmatch(r"[0-9a-fA-F]{64}", str(entry.get("sha256", ""))):
            raise ValueError(f"Incomplete immutable artifact: {entry}")
        path = Path(entry["path"])
        path = (base / path).resolve() if not path.is_absolute() else path.resolve()
        if str(path) in seen:
            raise ValueError(f"Duplicate immutable artifact: {path}")
        seen.add(str(path))
        if not path.is_file() or sha256(path) != entry["sha256"].lower():
            raise RuntimeError(f"Immutable analysis artifact changed: {path}")
        results.append({"path": str(path), "sha256": entry["sha256"].lower()})
    return sorted(results, key=lambda x: x["path"])


def _profile_length_cm(value: Any, field: str) -> float:
    """Parse the same explicit units accepted by the thesis adapter."""
    if isinstance(value, bool):
        raise ValueError(f"{field} must be a supported length, not boolean")
    match = re.fullmatch(
        r"\s*([0-9]+(?:\.[0-9]+)?)\s*(cm|mm|in|pt)?\s*",
        str(value).lower(),
    )
    if match is None:
        raise ValueError(
            f"{field} must be a supported length such as 21cm, 210mm, 8.5in, or 612pt"
        )
    return float(match.group(1)) * PROFILE_LENGTH_TO_CM[match.group(2) or "cm"]


def _manual_review_pages(
    page_count: int,
    touched_pages: list[int],
    adapter_pdf: dict[str, Any] | None = None,
) -> list[int]:
    """Select title/body samples and section-orientation boundaries."""
    if page_count < 1:
        return []
    pages = {1, min(2, page_count), min(3, page_count), page_count, *touched_pages}
    page_sizes = (adapter_pdf or {}).get("page_sizes", [])
    landscape = sorted({
        int(row["page"])
        for row in page_sizes
        if row.get("orientation") == "landscape"
        and 1 <= int(row.get("page", 0)) <= page_count
    })
    if landscape:
        group_start = landscape[0]
        previous = landscape[0]
        for page in landscape[1:]:
            if page != previous + 1:
                pages.update({
                    max(1, group_start - 1), group_start,
                    previous, min(page_count, previous + 1),
                })
                group_start = page
            previous = page
        pages.update({
            max(1, group_start - 1), group_start,
            previous, min(page_count, previous + 1),
        })
    return sorted(pages)


def file_state(path: Path) -> dict[str, Any]:
    return {
        "path": str(path),
        "exists": path.exists(),
        "size": path.stat().st_size if path.exists() else None,
        "sha256": sha256(path) if path.is_file() else None,
    }


def percentile(values: list[float], probability: float) -> float:
    ordered = sorted(values)
    if not ordered:
        return float("nan")
    position = (len(ordered) - 1) * probability
    lower = int(position)
    upper = min(lower + 1, len(ordered) - 1)
    fraction = position - lower
    return ordered[lower] + (ordered[upper] - ordered[lower]) * fraction


def _profile_path(value: str) -> Path:
    candidate = Path(value)
    if candidate.is_file():
        return candidate.resolve()
    bundled = PROFILE_DIR / f"{value}.json"
    if bundled.is_file():
        return bundled.resolve()
    raise FileNotFoundError(f"Unknown thesis profile: {value}")


def _benchmark_paths(base: Path, work_root: Path) -> tuple[Path, Path]:
    """Allocate archive/work directories without second-resolution collisions."""
    base.mkdir(parents=True, exist_ok=True)
    work_root.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%dT%H%M%S")
    session = Path(tempfile.mkdtemp(prefix=f"{stamp}_", dir=base))
    work_base = work_root / session.name
    work_base.mkdir(parents=True, exist_ok=False)
    return session, work_base


class Runner:
    def __init__(self) -> None:
        self.calls: list[dict[str, Any]] = []

    def run(
        self,
        command: list[str],
        *,
        check: bool = True,
        timeout: int = 600,
        env: dict[str, str] | None = None,
    ) -> subprocess.CompletedProcess[str]:
        is_word = any("osascript" in part for part in command) and any(
            "Microsoft Word" in part for part in command
        )
        with word_lock(" ".join(command)[:600]) if is_word else nullcontext():
            return self._run_locked(command, check=check, timeout=timeout, env=env)

    def _run_locked(self, command, *, check=True, timeout=600, env=None):
        if any("notebooklm" in part.lower() or part.lower() == "nlm" for part in command):
            raise RuntimeError("NotebookLM is disabled in the Word revision pipeline")
        started = time.perf_counter()
        merged_env = os.environ.copy()
        if env:
            merged_env.update(env)
        process = subprocess.Popen(
            [RTK, *command],
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            env=merged_env,
            start_new_session=True,
        )
        try:
            stdout, stderr = process.communicate(timeout=timeout)
        except subprocess.TimeoutExpired:
            try:
                os.killpg(process.pid, signal.SIGTERM)
            except ProcessLookupError:
                pass
            try:
                stdout, stderr = process.communicate(timeout=2)
            except subprocess.TimeoutExpired:
                try:
                    os.killpg(process.pid, signal.SIGKILL)
                except ProcessLookupError:
                    pass
                stdout, stderr = process.communicate()
            elapsed = time.perf_counter() - started
            self.calls.append({
                "command": command,
                "duration_seconds": round(elapsed, 4),
                "returncode": 124,
                "timed_out": True,
                "timeout_seconds": timeout,
            })
            raise RuntimeError(
                f"Command timed out after {timeout}s: {' '.join(command)}"
            )
        elapsed = time.perf_counter() - started
        completed = subprocess.CompletedProcess(
            [RTK, *command], process.returncode, stdout, stderr
        )
        self.calls.append({
            "command": command,
            "duration_seconds": round(elapsed, 4),
            "returncode": completed.returncode,
            "timed_out": False,
        })
        if check and completed.returncode:
            raise RuntimeError(
                f"Command failed ({completed.returncode}): {' '.join(command)}\n"
                f"STDOUT:\n{completed.stdout}\nSTDERR:\n{completed.stderr}"
            )
        return completed

    def office(self, *args: str, check: bool = True, timeout: int = 600):
        return self.run([OFFICECLI, *args], check=check, timeout=timeout)


class Pipeline:
    PHASES = ("toolchain", "audit", "layout", "revision", "word-postprocess", "verify")

    def __init__(self, job_path: Path, run_dir: Path | None = None) -> None:
        self.job_path = job_path.resolve()
        self.job = json.loads(self.job_path.read_text(encoding="utf-8"))
        self._normalize_job()
        self._resolve_job_paths()
        self._validate_job()
        self.run_dir = (run_dir or Path(self.job["run_dir"])).resolve()
        self.run_dir.mkdir(parents=True, exist_ok=True)
        self.cache_dir = Path(self.job["cache_dir"]).resolve()
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.manifest_path = self.run_dir / "run_manifest.json"
        self.manifest = self._load_manifest()
        self._command_index_base = len(self.manifest.get("commands", []))
        self.runner = Runner()
        self._saved_call_count = 0
        self.engine_hash = sha256(Path(__file__).resolve())
        self.adapter = self._load_adapter()
        if hasattr(self.adapter, "bind_runner"):
            self.adapter.bind_runner(self.runner)
        self.paths = self.adapter.configure(self.job, self.run_dir)

    def _normalize_job(self) -> None:
        """Normalize legacy jobs in memory without rewriting their source JSON."""
        if self.job.get("schema_version") == "1.0" and not self.job.get("scenario"):
            self.job["scenario"] = "manuscript_revision"
        scenario = self.job.get("scenario")
        if scenario == "thesis_format":
            self.job.setdefault("adapter", str(THESIS_ADAPTER))
            self.job.setdefault("content_source", self.job.get("source"))
            self.job.setdefault("structure_policy", "validate_only")
            self.job.setdefault("field_policy", "preserve")
            self.job.setdefault("table_policy", "report_only")
            self.job.setdefault("output", {"tracked_formatting": False})
        elif scenario == "manuscript_revision" and "publication_layout" in self.job:
            # Opt-in only: old manuscript jobs retain their historical behavior.
            self.job["publication_layout"] = normalize_publication_layout_policy(
                self.job["publication_layout"]
            )

    @property
    def scenario(self) -> str:
        return self.job.get("scenario", "manuscript_revision")

    def _implementation_hashes(self) -> dict[str, str]:
        paths = [Path(__file__).resolve(), Path(self.job["adapter"])]
        paths.append(SCRIPT_DIR / "word_runtime.py")
        paths.extend([SCRIPT_DIR / "journal_layout.py", SCRIPT_DIR / "manuscript_format.py", SCRIPT_DIR / "native_table_fit.py"])
        if hasattr(self.adapter, "dependency_paths"):
            paths.extend(Path(value) for value in self.adapter.dependency_paths())
        return {str(path): sha256(path) for path in paths if path.is_file()}

    def _resolve_job_paths(self) -> None:
        base = self.job_path.parent
        path_keys = {
            "source", "content_source", "adapter", "run_dir", "cache_dir",
            "analysis_manifest", "figure", "golden_accepted", "numeric_qa_script",
            "supplement", "profile",
        }
        for key in path_keys:
            value = self.job.get(key)
            if value and not Path(value).is_absolute():
                if key == "profile" and "/" not in value and "\\" not in value:
                    continue
                self.job[key] = str((base / value).resolve())
        if self.job.get("sources"):
            self.job["sources"] = [
                str(Path(value).resolve() if Path(value).is_absolute() else (base / value).resolve())
                for value in self.job["sources"]
            ]

    def _validate_job(self) -> None:
        required = ("schema_version", "job_id", "source", "adapter", "run_dir", "cache_dir")
        missing = [key for key in required if not self.job.get(key)]
        if missing:
            raise ValueError(f"Missing job fields: {missing}")
        non_strings = [
            key for key in required
            if not isinstance(self.job.get(key), str) or not self.job[key].strip()
        ]
        if non_strings:
            raise ValueError(f"Job fields must be non-empty strings: {non_strings}")
        if self.job.get("schema_version") not in {"1.0", "2.0"}:
            raise ValueError("schema_version must be 1.0 or 2.0")
        if self.scenario not in {"manuscript_revision", "thesis_format"}:
            raise ValueError("scenario must be manuscript_revision or thesis_format")
        if self.job.get("notebooklm") != "disabled":
            raise ValueError("job.notebooklm must be exactly 'disabled'; no enable branch exists")
        statistics_contract = self.job.get("statistics")
        if not isinstance(statistics_contract, dict):
            raise ValueError("job.statistics must be an object")
        if statistics_contract.get("allow_execution") is not False:
            raise ValueError("job.statistics.allow_execution must be exactly false")
        if statistics_contract.get("consume_manifest_only") is not True:
            raise ValueError("job.statistics.consume_manifest_only must be exactly true")
        required_paths = ["source", "adapter"]
        if self.scenario == "manuscript_revision":
            if not isinstance(self.job.get("content_source"), str) or not self.job["content_source"].strip():
                raise ValueError("manuscript_revision requires content_source")
            required_paths.append("content_source")
            if "publication_layout" in self.job:
                normalize_publication_layout_policy(self.job["publication_layout"])
        else:
            if "publication_layout" in self.job:
                raise ValueError("publication_layout is only valid for manuscript_revision")
            if self.job.get("scope") not in {"chapter", "chapter_batch", "full_thesis"}:
                raise ValueError("thesis_format requires scope=chapter, chapter_batch, or full_thesis")
            if not isinstance(self.job.get("profile"), str) or not self.job["profile"].strip():
                raise ValueError("thesis_format requires a versioned profile")
            sources = self.job.get("sources")
            if self.job.get("scope") == "chapter_batch":
                if not isinstance(sources, list) or not sources:
                    raise ValueError("chapter_batch requires a non-empty sources list")
                if any(not isinstance(value, str) or not value.strip() for value in sources):
                    raise ValueError("chapter_batch sources must be non-empty strings")
            elif sources is not None:
                raise ValueError("sources is only valid for thesis_format scope=chapter_batch")

            for key, allowed in THESIS_POLICY_VALUES.items():
                if self.job.get(key) not in allowed:
                    raise ValueError(f"{key} must be one of {sorted(allowed)}")

            output = self.job.get("output")
            if not isinstance(output, dict):
                raise ValueError("thesis_format output must be an object")
            if not isinstance(output.get("tracked_formatting"), bool):
                raise ValueError("output.tracked_formatting must be boolean")
            delivery_date = output.get("delivery_date")
            if delivery_date is not None and (
                not isinstance(delivery_date, str)
                or re.fullmatch(r"[0-9]{8}", delivery_date) is None
            ):
                raise ValueError("output.delivery_date must be YYYYMMDD")
            if delivery_date is not None:
                try:
                    datetime.strptime(delivery_date, "%Y%m%d")
                except ValueError as exc:
                    raise ValueError("output.delivery_date must be a real YYYYMMDD date") from exc

            style_map = self.job.get("style_map", {})
            if not isinstance(style_map, dict):
                raise ValueError("style_map must be an object")
            if any(
                not isinstance(role, str) or not role.strip()
                or not isinstance(selector, str) or not selector.strip()
                for role, selector in style_map.items()
            ):
                raise ValueError("style_map roles and selectors must be non-empty strings")
            for key in ("required_style_roles", "full_thesis_required_style_roles"):
                values = self.job.get(key)
                if values is not None and (
                    not isinstance(values, list)
                    or any(not isinstance(value, str) or not value.strip() for value in values)
                ):
                    raise ValueError(f"{key} must be an array of non-empty strings")

            profile_path = _profile_path(self.job["profile"])
            try:
                profile = json.loads(profile_path.read_text(encoding="utf-8"))
            except json.JSONDecodeError as exc:
                raise ValueError(f"Invalid thesis profile JSON: {profile_path}") from exc
            if not isinstance(profile, dict):
                raise ValueError("thesis profile must be a JSON object")
            missing_sections = sorted(PROFILE_REQUIRED_SECTIONS - profile.keys())
            if missing_sections:
                raise ValueError(f"Thesis profile missing required sections: {missing_sections}")
            expected_profile_id = (
                self.job["profile"]
                if "/" not in self.job["profile"] and "\\" not in self.job["profile"]
                else profile_path.stem
            )
            if profile.get("profile_id") != expected_profile_id:
                raise ValueError(
                    f"Thesis profile identity mismatch: expected {expected_profile_id!r}, "
                    f"found {profile.get('profile_id')!r}"
                )
            if not isinstance(profile.get("profile_version"), str) or re.fullmatch(
                r"[0-9]+\.[0-9]+\.[0-9]+", profile["profile_version"]
            ) is None:
                raise ValueError("thesis profile_version must be semantic version X.Y.Z")
            object_sections = ("document_props", "style_roles", "header", "footer", "page_numbering", "table")
            malformed_sections = [key for key in object_sections if not isinstance(profile.get(key), dict)]
            if malformed_sections:
                raise ValueError(f"Thesis profile sections must be objects: {malformed_sections}")
            missing_props = sorted(PROFILE_REQUIRED_DOCUMENT_PROPS - profile["document_props"].keys())
            if missing_props:
                raise ValueError(f"Thesis profile document_props missing: {missing_props}")
            geometry = {
                key: _profile_length_cm(
                    profile["document_props"][key], f"document_props.{key}"
                )
                for key in PROFILE_REQUIRED_DOCUMENT_PROPS
            }
            if geometry["pageWidth"] <= 0 or geometry["pageHeight"] <= 0:
                raise ValueError("Thesis profile page dimensions must be positive")
            if geometry["marginLeft"] + geometry["marginRight"] >= geometry["pageWidth"]:
                raise ValueError("Thesis profile horizontal margins must leave positive page width")
            if geometry["marginTop"] + geometry["marginBottom"] >= geometry["pageHeight"]:
                raise ValueError("Thesis profile vertical margins must leave positive page height")
            if not isinstance(profile.get("rules"), list) or not profile["rules"]:
                raise ValueError("thesis profile rules must be a non-empty array")
        for key in required_paths:
            if not Path(self.job[key]).exists():
                raise FileNotFoundError(self.job[key])
        for source in self.job.get("sources", []):
            if not Path(source).exists():
                raise FileNotFoundError(source)
        analysis_manifest = self.job.get("analysis_manifest")
        if analysis_manifest and not Path(analysis_manifest).exists():
            raise FileNotFoundError(analysis_manifest)
        if analysis_manifest and self.scenario == "thesis_format":
            raise ValueError("thesis_format does not consume statistical analysis manifests")
        if analysis_manifest:
            contract = json.loads(Path(analysis_manifest).read_text(encoding="utf-8"))
            if contract.get("statistics_execution_allowed_by_word_pipeline") is not False:
                raise ValueError("analysis_manifest must explicitly prohibit statistical execution")
            analysis_artifacts(contract, Path(analysis_manifest).parent)

    def _load_manifest(self) -> dict[str, Any]:
        if self.manifest_path.exists():
            return json.loads(self.manifest_path.read_text(encoding="utf-8"))
        return {
            "schema_version": "1.0",
            "job_id": self.job["job_id"],
            "job_path": str(self.job_path),
            "created_at": utc_now(),
            "updated_at": utc_now(),
            "notebooklm": "disabled",
            "scenario": self.scenario,
            "phases": {},
            "commands": [],
        }

    def _save_manifest(self) -> None:
        self.manifest["updated_at"] = utc_now()
        if len(self.runner.calls) > self._saved_call_count:
            self.manifest.setdefault("commands", []).extend(
                self.runner.calls[self._saved_call_count:]
            )
            self._saved_call_count = len(self.runner.calls)
        temp = self.manifest_path.with_suffix(".json.tmp")
        temp.write_text(json.dumps(self.manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        temp.replace(self.manifest_path)

    def _load_adapter(self):
        adapter_path = Path(self.job["adapter"])
        spec = importlib.util.spec_from_file_location(f"wordrev_adapter_{stable_hash(str(adapter_path))[:12]}", adapter_path)
        if spec is None or spec.loader is None:
            raise ImportError(adapter_path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module

    def _outputs_match(self, outputs: list[dict[str, Any]]) -> bool:
        for expected in outputs:
            path = Path(expected["path"])
            if not path.is_file() or sha256(path) != expected.get("sha256"):
                return False
        return True

    def _phase(
        self,
        name: str,
        fingerprint: str,
        action: Callable[[], list[Path]],
        *,
        resume: bool,
    ) -> dict[str, Any]:
        previous = self.manifest["phases"].get(name, {})
        if (
            resume
            and previous.get("status") == "complete"
            and previous.get("fingerprint") == fingerprint
            and self._outputs_match(previous.get("outputs", []))
        ):
            previous["cache_hit"] = True
            previous["last_checked_at"] = utc_now()
            self._save_manifest()
            return previous
        record = {
            "status": "running",
            "fingerprint": fingerprint,
            "cache_hit": False,
            "started_at": utc_now(),
            "command_start_index": self._command_index_base + len(self.runner.calls),
        }
        self.manifest["phases"][name] = record
        self._save_manifest()
        started = time.perf_counter()
        try:
            output_paths = action()
            record.update({
                "status": "complete",
                "finished_at": utc_now(),
                "duration_seconds": round(time.perf_counter() - started, 4),
                "outputs": [file_state(path) for path in output_paths],
                "command_end_index": self._command_index_base + len(self.runner.calls),
            })
        except Exception as exc:
            record.update({
                "status": "failed",
                "finished_at": utc_now(),
                "duration_seconds": round(time.perf_counter() - started, 4),
                "error": str(exc),
                "traceback": traceback.format_exc(),
                "command_end_index": self._command_index_base + len(self.runner.calls),
            })
            self._save_manifest()
            raise
        self._save_manifest()
        return record

    def _toolchain_state(self, *, force_refresh: bool = False) -> dict[str, Any]:
        cache = self.cache_dir / "toolchain.json"
        ttl = int(self.job.get("toolchain", {}).get("cache_ttl_hours", 24)) * 3600
        if cache.exists() and not force_refresh and time.time() - cache.stat().st_mtime < ttl:
            return json.loads(cache.read_text(encoding="utf-8"))
        office_version = self.runner.office("--version").stdout.strip()
        latest = self.runner.run(
            ["gh", "release", "view", "--repo", "iOfficeAI/OfficeCLI", "--json", "tagName,publishedAt"],
            check=False,
        )
        word = self.runner.run(
            ["osascript", "-e", 'tell application "Microsoft Word" to get version'], check=False
        )
        zotero = self.runner.run(["zotero-mcp", "version"], check=False)
        state = {
            "checked_at": utc_now(),
            "officecli_version": office_version,
            "officecli_latest_release": latest.stdout.strip(),
            "word_version": word.stdout.strip(),
            "zotero_mcp_version": zotero.stdout.strip(),
            "notebooklm": "disabled",
        }
        if self.job.get("toolchain", {}).get("require_latest", True):
            latest_tag = ""
            try:
                latest_tag = json.loads(latest.stdout).get("tagName", "").lstrip("v")
            except json.JSONDecodeError:
                pass
            if latest_tag and latest_tag != office_version.lstrip("v"):
                raise RuntimeError(
                    f"OfficeCLI {office_version} is not latest ({latest_tag}). "
                    "Upgrade outside document production, run smoke tests, then resume."
                )
        cache.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
        return state

    def toolchain(self, *, resume: bool, force_refresh: bool = False) -> dict[str, Any]:
        cache = self.cache_dir / "toolchain.json"
        ttl_seconds = int(self.job.get("toolchain", {}).get("cache_ttl_hours", 24)) * 3600
        fingerprint = stable_hash({
            "ttl": self.job.get("toolchain", {}).get("cache_ttl_hours", 24),
            "require_latest": self.job.get("toolchain", {}).get("require_latest", True),
            "refresh_window": int(time.time() // ttl_seconds),
            "force_refresh": force_refresh,
        })

        def action() -> list[Path]:
            state = self._toolchain_state(force_refresh=force_refresh)
            target = self.run_dir / "toolchain.json"
            target.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
            return [target, cache]

        return self._phase("toolchain", fingerprint, action, resume=resume)

    def _office_json(self, args: list[str], output: Path, *, check: bool = True) -> dict[str, Any]:
        process = self.runner.office(*args, check=check)
        output.write_text(process.stdout, encoding="utf-8")
        start = process.stdout.find("{")
        if start < 0:
            return {"raw": process.stdout, "stderr": process.stderr, "returncode": process.returncode}
        return json.loads(process.stdout[start:])

    def audit(self, *, resume: bool) -> dict[str, Any]:
        source = Path(self.job["source"])
        toolchain = json.loads((self.run_dir / "toolchain.json").read_text(encoding="utf-8"))
        audit_spec = self.job.get("audit", {})
        fingerprint = stable_hash({
            "source": sha256(source),
            "officecli": toolchain["officecli_version"],
            "scenario": self.scenario,
            "profile": self.job.get("profile"),
            "spec": audit_spec,
            "style_map": self.job.get("style_map", {}),
            "structure_policy": self.job.get("structure_policy"),
            "implementation": stable_hash(inspect.getsource(Pipeline.audit) + inspect.getsource(Pipeline._office_json)),
            "adapter_audit": inspect.getsource(self.adapter.augment_audit) if hasattr(self.adapter, "augment_audit") else None,
        })
        shared = self.cache_dir / "audit" / fingerprint
        target = self.run_dir / "audit"

        def action() -> list[Path]:
            if (shared / "audit.json").exists():
                if target.exists():
                    shutil.rmtree(target)
                shutil.copytree(shared, target)
                self.manifest["phases"]["audit"]["shared_cache_hit"] = True
                candidates = [
                    target / "audit.json", target / "source.pdf", target / "source_contact.png",
                    target / "format_diff.json", target / "format_diff.csv", target / "style_mapping_proposal.json",
                ]
                return [path for path in candidates if path.exists()]
            shared.mkdir(parents=True, exist_ok=True)
            raw = shared / "raw"
            raw.mkdir(exist_ok=True)
            results: dict[str, Any] = {
                "generated_at": utc_now(),
                "source": str(source),
                "source_sha256": sha256(source),
                "officecli_version": toolchain["officecli_version"],
                "officecli_latest_release": toolchain["officecli_latest_release"],
                "notebooklm": "disabled",
            }
            operations = {
                "document": ["get", str(source), "/", "--depth", "1", "--json"],
                "stats": ["view", str(source), "stats", "--json"],
                "outline": ["view", str(source), "outline", "--json"],
                "issues": ["view", str(source), "issues", "--json"],
                "styles": ["query", str(source), "style", "--json"],
                "tables": ["query", str(source), "table", "--json"],
                "images": ["query", str(source), "image", "--json"],
                "fields": ["query", str(source), "field", "--json"],
                "revisions": ["query", str(source), "revision", "--json"],
                "validation": ["validate", str(source), "--json"],
            }
            if self.scenario == "thesis_format":
                operations.update({
                    "sections": ["query", str(source), "section", "--json"],
                    "headers": ["query", str(source), "header", "--json"],
                    "footers": ["query", str(source), "footer", "--json"],
                    "footnotes": ["query", str(source), "footnote", "--json"],
                    "bookmarks": ["query", str(source), "bookmark", "--json"],
                })
            for name, args in operations.items():
                results[name] = self._office_json(
                    args, raw / f"{name}.json", check=self.scenario != "thesis_format"
                )
            results["selected_styles"] = {}
            for name, selector in audit_spec.get("styles", {}).items():
                results["selected_styles"][name] = self._office_json(
                    ["get", str(source), selector, "--depth", "5", "--json"],
                    raw / f"style_{name}.json",
                )
            results["anchors"] = {}
            for name, selector in audit_spec.get("anchors", {}).items():
                results["anchors"][name] = self._office_json(
                    ["get", str(source), selector, "--depth", "4", "--json"],
                    raw / f"anchor_{name}.json",
                )
            if audit_spec.get("render_source", True):
                pdf = shared / "source.pdf"
                contact = shared / "source_contact.png"
                self._export_pdf(source, pdf, "audit")
                self.runner.office(
                    "view", str(source), "screenshot", "--grid", "4", "--render", "html",
                    "-o", str(contact),
                )
                results["source_pdf"] = file_state(pdf)
                results["source_contact"] = file_state(contact)
            if hasattr(self.adapter, "augment_audit"):
                results["adapter_audit"] = self.adapter.augment_audit(
                    self.job, source, shared, results
                )
            (shared / "audit.json").write_text(
                json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8"
            )
            if target.exists():
                shutil.rmtree(target)
            shutil.copytree(shared, target)
            standard = [target / "audit.json", target / "source.pdf", target / "source_contact.png"]
            extras = [target / "format_diff.json", target / "format_diff.csv", target / "style_mapping_proposal.json"]
            return [path for path in [*standard, *extras] if path.exists()]

        return self._phase("audit", fingerprint, action, resume=resume)

    def _input_hashes(self) -> dict[str, str]:
        keys = ("source", "content_source", "adapter", "analysis_manifest", "figure")
        hashes = {}
        for key in keys:
            value = self.job.get(key)
            if value and Path(value).is_file():
                hashes[key] = sha256(Path(value))
        if self.job.get("analysis_manifest"):
            path = Path(self.job["analysis_manifest"])
            hashes["analysis_artifacts"] = stable_hash(analysis_artifacts(json.loads(path.read_text()), path.parent))
        return hashes

    def _export_pdf(self, source: Path, output: Path, label: str) -> None:
        staging_value = self.job.get("word_pdf_staging_dir")
        if not staging_value:
            self.adapter.export_pdf(source, output, print_markup=False)
            return
        staging_dir = Path(staging_value).resolve()
        staging_dir.mkdir(parents=True, exist_ok=True)
        safe_job_id = re.sub(r"[^A-Za-z0-9_.-]+", "-", self.job["job_id"])
        nonce = uuid.uuid4().hex[:10]
        staging_source = staging_dir / f"wordrev_{safe_job_id}_{label}_{nonce}.docx"
        staging_pdf = staging_dir / f"wordrev_{safe_job_id}_{label}_{nonce}.pdf"
        staging_source.unlink(missing_ok=True)
        staging_pdf.unlink(missing_ok=True)
        succeeded = False
        try:
            if hasattr(self.adapter, "prepare_word_staging"):
                # Thesis files may request field refresh on open. Work on a
                # disposable copy so the adapter can disable that modal trigger
                # before Microsoft Word performs authoritative pagination.
                shutil.copy2(source, staging_source)
                self.adapter.prepare_word_staging(staging_source, purpose=label)
                export_source = staging_source
            else:
                # Preserve the proven manuscript adapter path for schema 1.0
                # and existing v7.1 jobs.
                export_source = source
            self.adapter.export_pdf(export_source, staging_pdf, print_markup=False)
            output.parent.mkdir(parents=True, exist_ok=True)
            output.unlink(missing_ok=True)
            shutil.move(staging_pdf, output)
            succeeded = True
        finally:
            # Failed owned operations retain their exact recovery artifacts.
            # Never unlink an input still open in Word after a modal/timeout.
            if succeeded:
                staging_source.unlink(missing_ok=True)
                staging_pdf.unlink(missing_ok=True)

    def layout(self, *, resume: bool) -> dict[str, Any]:
        fingerprint = stable_hash({
            "inputs": self._input_hashes(),
            "audit": sha256(self.run_dir / "audit" / "audit.json"),
            "layout": self.job.get("layout", {}),
            "publication_layout": self.job.get("publication_layout"),
            "format_contract": self.job.get("format_contract"),
            "citations": self.job.get("citations", []),
            "scenario_contract": {
                "scenario": self.scenario,
                "scope": self.job.get("scope"),
                "profile": self.job.get("profile"),
                "thesis_title": self.job.get("thesis_title"),
                "style_map": self.job.get("style_map", {}),
                "section_map": self.job.get("section_map", {}),
                "structure_policy": self.job.get("structure_policy"),
                "field_policy": self.job.get("field_policy"),
                "field_operations": self.job.get("field_operations", []),
                "table_policy": self.job.get("table_policy"),
                "table_indices": self.job.get("table_indices", []),
            },
            "implementation": self._implementation_hashes(),
        })
        cache = self.cache_dir / "layout" / fingerprint
        prepared = Path(self.paths["prepared"])
        metadata_path = Path(self.paths["metadata"])
        layout_pdf = Path(self.paths["layout_pdf"])
        checkpoint = metadata_path.parent / "layout_content_checkpoint.json"

        def layout_sidecars() -> list[Path]:
            if self.scenario != "thesis_format":
                return []
            return [
                Path(self.paths[key])
                for key in ("semantic_before", "command_count")
                if self.paths.get(key)
            ]

        def action() -> list[Path]:
            cached_prepared = cache / "prepared.docx"
            cached_metadata = cache / "metadata.json"
            cached_pdf = cache / "layout.pdf"
            if cached_prepared.exists() and cached_metadata.exists() and cached_pdf.exists():
                prepared.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(cached_prepared, prepared)
                shutil.copy2(cached_metadata, metadata_path)
                shutil.copy2(cached_pdf, layout_pdf)
                metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
                if hasattr(self.adapter, "restore_layout_sidecars"):
                    self.adapter.restore_layout_sidecars(metadata, self.paths)
                self.manifest["phases"]["layout"]["shared_cache_hit"] = True
                return [prepared, metadata_path, layout_pdf, *layout_sidecars()]
            previous = json.loads(checkpoint.read_text()) if checkpoint.exists() else {}
            content_ready = (
                previous.get("fingerprint") == fingerprint
                and prepared.is_file() and metadata_path.is_file()
                and previous.get("prepared_sha256") == sha256(prepared)
                and previous.get("metadata_sha256") == sha256(metadata_path)
            )
            if not content_ready:
                metadata = self.adapter.build_layout(self.job, self.paths)
                if self.scenario == "manuscript_revision" and self.job.get("publication_layout"):
                    metadata["native_table_fit"] = native_fit(self.runner.run, prepared, self.job["publication_layout"])
                metadata_path.write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
                checkpoint.write_text(json.dumps({
                    "fingerprint": fingerprint, "prepared_sha256": sha256(prepared),
                    "metadata_sha256": sha256(metadata_path),
                }, indent=2))
            self.manifest["phases"]["layout"]["content_checkpoint_hit"] = content_ready
            self._export_pdf(prepared, layout_pdf, "layout")
            cache.mkdir(parents=True, exist_ok=True)
            shutil.copy2(prepared, cached_prepared)
            shutil.copy2(metadata_path, cached_metadata)
            shutil.copy2(layout_pdf, cached_pdf)
            return [prepared, metadata_path, layout_pdf, *layout_sidecars()]

        return self._phase("layout", fingerprint, action, resume=resume)

    def revision(self, *, resume: bool) -> dict[str, Any]:
        if self.scenario == "thesis_format" and not self.job.get("output", {}).get("tracked_formatting", False):
            record = {
                "status": "skipped",
                "reason": "thesis clean-output mode",
                "cache_hit": True,
                "finished_at": utc_now(),
                "outputs": [],
            }
            self.manifest["phases"]["revision"] = record
            self._save_manifest()
            return record
        metadata_path = Path(self.paths["metadata"])
        fingerprint = stable_hash({
            "prepared": sha256(Path(self.paths["prepared"])),
            "metadata": sha256(metadata_path),
            "author": self.job.get("revision_author", "Codex"),
            "revision_date": self.job.get("revision_date"),
            "implementation": self._implementation_hashes(),
        })

        def action() -> list[Path]:
            metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
            self.adapter.build_tracked(self.job, self.paths, metadata)
            if self.scenario == "manuscript_revision" and self.job.get("publication_layout"):
                rows = transfer_geometry(self.paths["prepared"], self.paths["tracked"],
                    self.job["publication_layout"], self.job.get("revision_author","Codex"), self.job["revision_date"])
                (self.run_dir/"evidence"/"native_table_fit_revisions.json").write_text(json.dumps(rows,indent=2))
            return [Path(self.paths["tracked"]), Path(self.paths["batch"]), Path(self.paths["command_count"])]

        return self._phase("revision", fingerprint, action, resume=resume)

    def word_postprocess(self, *, resume: bool) -> dict[str, Any]:
        source_key = "tracked" if self.scenario == "manuscript_revision" else "prepared"
        fingerprint = stable_hash({
            source_key: sha256(Path(self.paths[source_key])),
            "implementation": self._implementation_hashes(),
            "citations": self.job.get("citations", []),
            "field_policy": self.job.get("field_policy"),
            "scope": self.job.get("scope"),
            "output": self.job.get("output", {}),
        })

        def action() -> list[Path]:
            self.adapter.postprocess(self.job, self.paths)
            if hasattr(self.adapter, "phase_outputs"):
                values = self.adapter.phase_outputs("word-postprocess", self.job, self.paths)
                return [Path(value) for value in values]
            return [Path(self.paths["accepted"]), Path(self.paths["rejected"])]

        return self._phase("word-postprocess", fingerprint, action, resume=resume)

    def build(self, *, resume: bool) -> None:
        self.toolchain(resume=resume)
        self.audit(resume=resume)
        self.layout(resume=resume)
        self.revision(resume=resume)
        self.word_postprocess(resume=resume)

    def _numeric_qa(self) -> dict[str, Any]:
        script = self.job.get("numeric_qa_script")
        if not script:
            return {"skipped": True}
        output = self.run_dir / "numeric_qa"
        output.mkdir(exist_ok=True)
        python = self.job.get(
            "python",
            "/Users/lzhs/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3",
        )
        env = {
            "V71_MAIN_DOC": self.paths["accepted"],
            "V71_EVIDENCE_DIR": str(output),
        }
        process = self.runner.run([python, script], env=env)
        result_path = output / "delivery_qa.json"
        result = json.loads(result_path.read_text(encoding="utf-8"))
        if not result.get("all_pass"):
            raise RuntimeError("Numeric QA did not pass")
        expected = self.job.get("regression", {}).get("numeric_checks_expected")
        if expected is not None and result.get("n_checks") != expected:
            raise RuntimeError(
                f"Numeric QA count changed: expected {expected}, got {result.get('n_checks')}"
            )
        return {"stdout": process.stdout.strip(), "result": result, "path": str(result_path)}

    def _pdf_checks(self, pdf: Path) -> dict[str, Any]:
        pdfinfo = self.runner.run(["pdfinfo", str(pdf)])
        page_count = 0
        for line in pdfinfo.stdout.splitlines():
            if line.startswith("Pages:"):
                page_count = int(line.split(":", 1)[1].strip())
        text_dir = self.run_dir / "pdf_qa"
        text_dir.mkdir(exist_ok=True)
        page_metrics = []
        render_prefix = text_dir / "page"
        self.runner.run(["pdftoppm", "-r", "96", "-png", str(pdf), str(render_prefix)], timeout=600)
        try:
            from PIL import Image, ImageStat
        except ImportError as exc:
            raise RuntimeError("Pillow is required for PDF page checks") from exc
        page_images = sorted(text_dir.glob("page-*.png"))
        for index, image_path in enumerate(page_images, start=1):
            with Image.open(image_path).convert("L") as image:
                stat = ImageStat.Stat(image)
                mean = stat.mean[0]
                extrema = stat.extrema[0]
                sample = image.resize((120, 160))
                pixels = sample.get_flattened_data() if hasattr(sample, "get_flattened_data") else sample.getdata()
                nonwhite = sum(1 for value in pixels if value < 248)
                page_metrics.append({
                    "page": index,
                    "mean_gray": round(mean, 3),
                    "min_gray": extrema[0],
                    "sample_nonwhite_pixels": nonwhite,
                    "nonblank": nonwhite >= 8,
                })
        expected_pages = self.job.get("visual_qa", {}).get("expected_page_range", [])
        if expected_pages and not (expected_pages[0] <= page_count <= expected_pages[1]):
            raise RuntimeError(f"Unexpected page count: {page_count}")
        blank = [row["page"] for row in page_metrics if not row["nonblank"]]
        allowed_blank = set(self.job.get("visual_qa", {}).get("allowed_blank_pages", []))
        unexpected_blank = [page for page in blank if page not in allowed_blank]
        if unexpected_blank:
            raise RuntimeError(f"Unexpected blank PDF pages detected: {unexpected_blank}")
        return {
            "page_count": page_count,
            "pages": page_metrics,
            "blank_pages": blank,
            "allowed_blank_pages": sorted(allowed_blank),
            "unexpected_blank_pages": unexpected_blank,
        }

    def _touched_contact(
        self,
        pdf: Path,
        page_count: int,
        adapter_pdf: dict[str, Any] | None = None,
    ) -> Path:
        touched = self.job.get("visual_qa", {}).get("touched_pages", [])
        invalid_pages = [page for page in touched if page > page_count]
        if invalid_pages:
            raise ValueError(
                f"visual_qa.touched_pages are 1-indexed and must not exceed {page_count}: "
                f"{invalid_pages}"
            )
        pages = _manual_review_pages(page_count, touched, adapter_pdf)
        output_dir = self.run_dir / "touched_pages"
        output_dir.mkdir(exist_ok=True)
        if not pages:
            return output_dir / "none.txt"
        try:
            from PIL import Image, ImageOps, ImageDraw
        except ImportError as exc:
            raise RuntimeError("Pillow is required for contact sheets") from exc
        images = []
        for page in pages:
            prefix = output_dir / f"page_{page}"
            self.runner.run([
                "pdftoppm", "-f", str(page), "-l", str(page), "-r", "130", "-png",
                str(pdf), str(prefix),
            ])
            candidate = next(output_dir.glob(f"page_{page}-*.png"))
            image = Image.open(candidate).convert("RGB")
            image.thumbnail((620, 880))
            canvas = Image.new("RGB", (640, 920), "white")
            canvas.paste(image, ((640 - image.width) // 2, 30))
            ImageDraw.Draw(canvas).text((12, 8), f"Page {page}", fill="black")
            images.append(canvas)
        contact = output_dir / "manual_review_pages_contact.png"
        columns = min(4, len(images))
        rows = (len(images) + columns - 1) // columns
        sheet = Image.new("RGB", (640 * columns, 920 * rows), "white")
        for index, image in enumerate(images):
            sheet.paste(
                ImageOps.expand(image, border=1, fill="gray"),
                (640 * (index % columns), 920 * (index // columns)),
            )
        sheet.save(contact)
        return contact

    def verify(self, *, resume: bool) -> dict[str, Any]:
        accepted = Path(self.paths["accepted"])
        tracked = Path(self.paths["tracked"]) if self.paths.get("tracked") else None
        rejected = Path(self.paths["rejected"]) if self.paths.get("rejected") else None
        input_hashes = {"accepted": sha256(accepted)}
        if tracked and tracked.is_file():
            input_hashes["tracked"] = sha256(tracked)
        if rejected and rejected.is_file():
            input_hashes["rejected"] = sha256(rejected)
        fingerprint = stable_hash({
            "documents": input_hashes,
            "numeric_qa_script": sha256(Path(self.job["numeric_qa_script"])) if self.job.get("numeric_qa_script") else None,
            "analysis_inputs": self._input_hashes(),
            "audit": sha256(self.run_dir / "audit" / "audit.json"),
            "qa": self.job.get("regression", {}),
            "visual": self.job.get("visual_qa", {}),
            "publication_layout": self.job.get("publication_layout"),
            "format_contract": self.job.get("format_contract"),
            "implementation": self._implementation_hashes(),
        })
        report_path = self.run_dir / "verification.json"

        def action() -> list[Path]:
            audit = json.loads((self.run_dir / "audit" / "audit.json").read_text(encoding="utf-8"))
            command_path = Path(self.paths["command_count"]) if self.paths.get("command_count") else None
            command_count = int(command_path.read_text(encoding="utf-8")) if command_path and command_path.exists() else 0
            adapter_qa = self.adapter.verify(self.job, self.paths, audit, command_count)
            accepted_pdf = Path(self.paths["accepted_pdf"])
            self._export_pdf(accepted, accepted_pdf, "accepted")
            # Shared final-file gates cannot be replaced by a project's narrower
            # numeric or caption-presence checks. Check AFTER Word round-trip.
            shared_docx_qa = {}
            if self.scenario == "manuscript_revision":
                from manuscript_format import audit_native_clean
                shared_docx_qa['accepted_native_revisions'] = audit_native_clean(accepted)
                if rejected:
                    shared_docx_qa['rejected_native_revisions'] = audit_native_clean(rejected)
            if self.scenario == "manuscript_revision" and "publication_layout" in self.job:
                from journal_layout import audit_table_layout
                shared_docx_qa["table_layout"] = audit_table_layout(
                    accepted, self.job["publication_layout"],
                    math_line_rule=self.job.get("format_contract", {}).get("table", {}).get("math_line_rule", "exact"),
                )
            if self.job.get("format_contract"):
                from manuscript_format import audit_contract
                for label, path in [("accepted", accepted), ("tracked", tracked)]:
                    if path:
                        shared_docx_qa[label+"_format"] = audit_contract(path,self.job["format_contract"],self.job.get("publication_layout"))
            shared_path = self.run_dir / "final_format_qa.json"
            shared_path.write_text(json.dumps(shared_docx_qa,ensure_ascii=False,indent=2))
            # The files are frozen at this point. These checks are read-only and
            # may safely run in parallel; OfficeCLI/Word writes never overlap.
            with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
                numeric_future = executor.submit(self._numeric_qa)
                pdf_future = executor.submit(self._pdf_checks, accepted_pdf)
                numeric = numeric_future.result()
                pdf = pdf_future.result()
            adapter_pdf = (
                self.adapter.verify_pdf(self.job, self.paths, accepted_pdf, pdf)
                if hasattr(self.adapter, "verify_pdf") else {"all_pass": True}
            )
            if "publication_layout" in self.job:
                publication_qa = adapter_pdf.get("publication_layout_qa", {})
                if not publication_qa.get("all_pass"):
                    raise RuntimeError(
                        "publication_layout requires adapter.verify_pdf with passing publication_layout_qa"
                    )
            contact = self._touched_contact(
                accepted_pdf, pdf["page_count"], adapter_pdf
            )
            result = {
                "generated_at": utc_now(),
                "all_pass": (
                    bool(adapter_qa.get("all_pass", True))
                    and bool(numeric.get("result", {}).get("all_pass", True))
                    and bool(adapter_pdf.get("all_pass", True))
                    and all(bool(q.get("all_pass")) for q in shared_docx_qa.values())
                ),
                "adapter_qa": adapter_qa,
                "shared_final_docx_qa": shared_docx_qa,
                "adapter_pdf_qa": adapter_pdf,
                "numeric_qa": numeric,
                "pdf_qa": pdf,
                "touched_pages_contact": str(contact),
                "notebooklm": "disabled",
                "hashes": {
                    key: sha256(Path(value))
                    for key, value in self.paths.items()
                    if key in {"tracked", "accepted", "rejected"} and Path(value).is_file()
                },
            }
            report_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
            if not result["all_pass"]:
                raise RuntimeError("Verification did not pass")
            return [report_path, accepted_pdf, contact, shared_path]

        return self._phase("verify", fingerprint, action, resume=resume)


def load_job_run_dir(job_path: Path) -> Path:
    data = json.loads(job_path.read_text(encoding="utf-8"))
    value = Path(data["run_dir"])
    return value if value.is_absolute() else (job_path.parent / value).resolve()


def _slug(value: str) -> str:
    cleaned = re.sub(r"[^0-9A-Za-z._-]+", "-", value).strip("-._")
    return cleaned[:80] or "document"


def _validated_job_document(job_path: Path) -> dict[str, Any]:
    """Normalize and validate a job without creating run/cache directories."""
    probe = Pipeline.__new__(Pipeline)
    probe.job_path = job_path.resolve()
    probe.job = json.loads(probe.job_path.read_text(encoding="utf-8"))
    probe._normalize_job()
    probe._resolve_job_paths()
    probe._validate_job()
    return probe.job


def init_job(args: argparse.Namespace) -> Path:
    sources = [Path(value).expanduser().resolve() for value in args.source]
    for source in sources:
        if not source.is_file():
            raise FileNotFoundError(source)
    scenario = args.scenario.replace("-", "_")
    scope = args.scope.replace("-", "_") if args.scope else None
    primary = sources[0]
    run_dir = Path(args.run_dir).expanduser().resolve() if args.run_dir else primary.parent / f"{primary.stem}_wordrev"
    cache_dir = Path(args.cache_dir).expanduser().resolve() if args.cache_dir else Path.home() / ".cache" / "wordrev"
    staging_dir = (
        Path(args.word_pdf_staging_dir).expanduser().resolve()
        if args.word_pdf_staging_dir
        else Path(os.environ.get("WORDREV_STAGING_DIR", Path.home() / "Documents" / "wordrev-pdf-staging"))
    )
    job: dict[str, Any] = {
        "$schema": str(SKILL_DIR / "references" / "job.schema.json"),
        "schema_version": "2.0",
        "job_id": args.job_id or f"{_slug(primary.stem)}-{scenario.replace('_', '-')}",
        "scenario": scenario,
        "source": str(primary),
        "run_dir": str(run_dir),
        "cache_dir": str(cache_dir),
        "benchmark_dir": str(run_dir.parent / f"{run_dir.name}_benchmarks"),
        "word_pdf_staging_dir": str(staging_dir),
        "notebooklm": "disabled",
        "statistics": {"allow_execution": False, "consume_manifest_only": True},
        "toolchain": {"cache_ttl_hours": 24, "require_latest": True, "auto_upgrade": False},
        "audit": {"render_source": True, "styles": {}, "anchors": {}},
        "visual_qa": {"touched_pages": []},
    }
    if scenario == "thesis_format":
        if not scope:
            raise ValueError("thesis-format init requires --scope")
        style_map = {}
        for entry in args.style or []:
            if "=" not in entry:
                raise ValueError(f"Invalid --style {entry!r}; expected role=/styles/ID")
            role, selector = entry.split("=", 1)
            style_map[role] = selector
        job.update({
            "scope": scope,
            "sources": [str(path) for path in sources] if scope == "chapter_batch" else None,
            "profile": args.profile or "ruc-doctoral-2026",
            "adapter": str(THESIS_ADAPTER),
            "thesis_title": args.thesis_title or "",
            "style_map": style_map,
            "section_map": {},
            "structure_policy": "validate_only",
            "field_policy": "preserve",
            "table_policy": "report_only",
            "output": {
                "tracked_formatting": False,
                "delivery_date": datetime.now().strftime("%Y%m%d"),
            },
        })
        if job["sources"] is None:
            job.pop("sources")
    else:
        if len(sources) != 1 or not args.content_source or not args.adapter:
            raise ValueError("manuscript-revision init requires one --source, --content-source, and --adapter")
        job.update({
            "content_source": str(Path(args.content_source).expanduser().resolve()),
            "adapter": str(Path(args.adapter).expanduser().resolve()),
            "revision_author": args.revision_author,
            "citations": [],
            "publication_layout": default_publication_layout_policy(),
        })
    output = Path(args.out).expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        raise FileExistsError(f"Refusing to overwrite existing job: {output}")
    output.write_text(json.dumps(job, ensure_ascii=False, indent=2), encoding="utf-8")
    return output


def _chapter_jobs(job_path: Path) -> list[tuple[Path, Path]]:
    raw = json.loads(job_path.read_text(encoding="utf-8"))
    if raw.get("scenario") != "thesis_format" or raw.get("scope") != "chapter_batch":
        return []
    job = _validated_job_document(job_path)
    if job.get("scenario") != "thesis_format" or job.get("scope") != "chapter_batch":
        return []
    base = Path(job["run_dir"])
    if not base.is_absolute():
        base = (job_path.parent / base).resolve()
    definitions = base / "chapter_jobs"
    definitions.mkdir(parents=True, exist_ok=True)
    rows = []
    for index, value in enumerate(job.get("sources", []), start=1):
        source = Path(value)
        if not source.is_absolute():
            source = (job_path.parent / source).resolve()
        child = dict(job)
        child.pop("sources", None)
        child["scope"] = "chapter"
        child["source"] = str(source)
        child["job_id"] = f"{job['job_id']}-{index:02d}-{_slug(source.stem)}"
        child_run = base / f"{index:02d}_{_slug(source.stem)}"
        child["run_dir"] = str(child_run)
        child_path = definitions / f"{index:02d}_{_slug(source.stem)}.json"
        child_path.write_text(json.dumps(child, ensure_ascii=False, indent=2), encoding="utf-8")
        rows.append((child_path, child_run))
    return rows


def run_chapter_batch(job_path: Path, command: str, *, resume: bool) -> dict[str, Any]:
    rows = []
    for child_path, child_run in _chapter_jobs(job_path):
        pipeline = Pipeline(child_path, child_run)
        if command == "audit":
            pipeline.toolchain(resume=resume)
            pipeline.audit(resume=resume)
        elif command == "layout":
            pipeline.toolchain(resume=resume)
            pipeline.audit(resume=resume)
            pipeline.layout(resume=resume)
        elif command == "build":
            pipeline.build(resume=resume)
        elif command == "verify":
            pipeline.verify(resume=resume)
        rows.append({"job": str(child_path), "run_dir": str(child_run), "command": command, "status": "complete"})
    base = load_job_run_dir(job_path)
    base.mkdir(parents=True, exist_ok=True)
    summary = {"scenario": "thesis_format", "scope": "chapter_batch", "command": command, "chapters": rows}
    (base / "chapter_batch_manifest.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    return summary


def benchmark(job_path: Path, cold: int, warm: int) -> dict[str, Any]:
    if isinstance(cold, bool) or not isinstance(cold, int) or cold <= 0:
        raise ValueError("benchmark cold sample count must be a positive integer")
    if isinstance(warm, bool) or not isinstance(warm, int) or warm <= 0:
        raise ValueError("benchmark warm sample count must be a positive integer")
    job = _validated_job_document(job_path)
    if job.get("scenario") == "thesis_format" and job.get("scope") == "chapter_batch":
        raise ValueError("Benchmark chapter_batch child jobs individually; no chapter merging is performed")
    configured_base = job.get("benchmark_dir")
    if configured_base:
        base = Path(configured_base)
        if not base.is_absolute():
            base = (job_path.parent / base).resolve()
    else:
        run_dir = load_job_run_dir(job_path)
        base = run_dir.parent / f"{run_dir.name}_benchmarks"
    work_root = Path(job.get("benchmark_work_dir", "/tmp/wordrev"))
    session, work_base = _benchmark_paths(base, work_root)
    rows = []
    last_run = None
    for index in range(1, cold + 1):
        run_dir = work_base / f"c{index}"
        started = time.perf_counter()
        pipeline = Pipeline(job_path, run_dir)
        pipeline.build(resume=False)
        pipeline.verify(resume=False)
        elapsed = time.perf_counter() - started
        rows.append({"kind": "cold", "iteration": index, "seconds": elapsed, "run_dir": str(run_dir)})
        last_run = run_dir
    if last_run is None:
        last_run = load_job_run_dir(job_path)
        pipeline = Pipeline(job_path, last_run)
        pipeline.build(resume=True)
        pipeline.verify(resume=True)
    for index in range(1, warm + 1):
        started = time.perf_counter()
        pipeline = Pipeline(job_path, last_run)
        pipeline.build(resume=True)
        pipeline.verify(resume=True)
        elapsed = time.perf_counter() - started
        rows.append({"kind": "warm", "iteration": index, "seconds": elapsed, "run_dir": str(last_run)})
    summary = {}
    for kind in ("cold", "warm"):
        values = [row["seconds"] for row in rows if row["kind"] == kind]
        summary[kind] = {
            "n": len(values),
            "median_seconds": statistics.median(values) if values else None,
            "p95_seconds": percentile(values, 0.95) if values else None,
        }
    thresholds = job.get("performance_thresholds", {})
    summary["thresholds"] = {
        "warm_resume_median_le_seconds": thresholds.get("warm_resume_median_seconds", 20),
        "cold_cached_median_le_seconds": thresholds.get("cold_cached_median_seconds", 120),
        "cold_cached_p95_le_seconds": thresholds.get("cold_cached_p95_seconds", 180),
    }
    summary["pass"] = (
        (summary["warm"]["median_seconds"] or 0) <= summary["thresholds"]["warm_resume_median_le_seconds"]
        and (summary["cold"]["median_seconds"] or 0) <= summary["thresholds"]["cold_cached_median_le_seconds"]
        and (summary["cold"]["p95_seconds"] or 0) <= summary["thresholds"]["cold_cached_p95_le_seconds"]
    )
    payload = {
        "generated_at": utc_now(),
        "job": str(job_path),
        "benchmark_archive": str(session),
        "benchmark_work_dir": str(work_base),
        "rows": rows,
        "summary": summary,
    }
    json_path = session / "benchmark.json"
    csv_path = session / "benchmark.csv"
    md_path = session / "benchmark.md"
    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    with csv_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["kind", "iteration", "seconds", "run_dir"])
        writer.writeheader()
        writer.writerows(rows)
    md_path.write_text(
        "# Word revision pipeline benchmark\n\n"
        f"- Cold cached rebuild median: `{summary['cold']['median_seconds']:.2f}s`; "
        f"P95: `{summary['cold']['p95_seconds']:.2f}s`.\n"
        f"- Warm no-change resume median: `{summary['warm']['median_seconds']:.2f}s`; "
        f"P95: `{summary['warm']['p95_seconds']:.2f}s`.\n"
        f"- Thresholds passed: `{summary['pass']}`.\n"
        "- Speedup is not claimed unless golden regression and timing thresholds both pass.\n",
        encoding="utf-8",
    )
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return payload


def parser() -> argparse.ArgumentParser:
    root = argparse.ArgumentParser(prog="wordrev")
    sub = root.add_subparsers(dest="command", required=True)
    init = sub.add_parser("init")
    init.add_argument("--scenario", required=True, choices=("manuscript-revision", "thesis-format"))
    init.add_argument("--scope", choices=("chapter", "chapter-batch", "full-thesis"))
    init.add_argument("--source", required=True, action="append")
    init.add_argument("--content-source")
    init.add_argument("--adapter")
    init.add_argument("--profile", default="ruc-doctoral-2026")
    init.add_argument("--thesis-title")
    init.add_argument("--style", action="append", help="Semantic style mapping, e.g. body=/styles/Normal")
    init.add_argument("--job-id")
    init.add_argument("--revision-author", default="Codex")
    init.add_argument("--run-dir")
    init.add_argument("--cache-dir")
    init.add_argument("--word-pdf-staging-dir")
    init.add_argument("--out", required=True)
    for name in ("audit", "layout", "build"):
        command = sub.add_parser(name)
        command.add_argument("--job", required=True, type=Path)
        command.add_argument("--run-dir", type=Path)
        command.add_argument("--resume", action="store_true")
    verify = sub.add_parser("verify")
    verify.add_argument("--run-dir", required=True, type=Path)
    verify.add_argument("--job", type=Path)
    verify.add_argument("--resume", action="store_true")
    bench = sub.add_parser("benchmark")
    bench.add_argument("--job", required=True, type=Path)
    bench.add_argument("--cold", type=int, default=3)
    bench.add_argument("--warm", type=int, default=5)
    return root


def main() -> None:
    args = parser().parse_args()
    if args.command == "init":
        output = init_job(args)
        print(json.dumps({"created": str(output)}, ensure_ascii=False, indent=2))
        return
    if args.command == "benchmark":
        benchmark(args.job.resolve(), args.cold, args.warm)
        return
    if args.command in {"audit", "layout", "build"}:
        batch_jobs = _chapter_jobs(args.job.resolve())
        if batch_jobs:
            print(json.dumps(run_chapter_batch(args.job.resolve(), args.command, resume=args.resume), ensure_ascii=False, indent=2))
            return
    if args.command == "verify":
        job_path = args.job
        if job_path is None:
            manifest = json.loads((args.run_dir / "run_manifest.json").read_text(encoding="utf-8"))
            job_path = Path(manifest["job_path"])
        if _chapter_jobs(job_path.resolve()):
            print(json.dumps(run_chapter_batch(job_path.resolve(), "verify", resume=args.resume), ensure_ascii=False, indent=2))
            return
        pipeline = Pipeline(job_path.resolve(), args.run_dir)
        pipeline.verify(resume=args.resume)
        print(json.dumps(pipeline.manifest["phases"]["verify"], ensure_ascii=False, indent=2))
        return
    pipeline = Pipeline(args.job.resolve(), args.run_dir)
    if args.command == "audit":
        pipeline.toolchain(resume=args.resume)
        result = pipeline.audit(resume=args.resume)
    elif args.command == "layout":
        pipeline.toolchain(resume=args.resume)
        pipeline.audit(resume=args.resume)
        result = pipeline.layout(resume=args.resume)
    else:
        pipeline.build(resume=args.resume)
        result = pipeline.manifest["phases"]["word-postprocess"]
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
