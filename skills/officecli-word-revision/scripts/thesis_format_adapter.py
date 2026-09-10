"""Generic institution-profile adapter for doctoral thesis formatting."""

from __future__ import annotations

import csv
import hashlib
import json
import os
import posixpath
import re
import shutil
import subprocess
import tempfile
import time
import uuid
import zipfile
from contextlib import contextmanager
from pathlib import Path, PurePosixPath
from typing import Any

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Mm, Pt
from lxml import etree

try:
    import fcntl
    HAS_FCNTL = True
except ImportError:
    HAS_FCNTL = False


SKILL_DIR = Path(__file__).resolve().parent.parent
PROFILE_DIR = SKILL_DIR / "references" / "profiles"
RUNNER = None
WORD_LOCK_TIMEOUT_SECONDS = 300.0
SEMANTIC_INVARIANT_KEYS = (
    "body_text_sha256",
    "footnote_text_sha256",
    "media_payload_sha256",
    "media_usage_sha256",
    "equation_xml_sha256",
    "tables",
    "equations",
    "drawings",
    "zotero_fields",
    "zotero_fields_sha256",
)


def dependency_paths():
    return sorted(PROFILE_DIR.glob("*.json"))


def bind_runner(runner) -> None:
    global RUNNER
    RUNNER = runner


def run(command, *, check=True, timeout=300):
    if RUNNER is not None:
        return RUNNER.run(command, check=check, timeout=timeout)
    process = subprocess.run(command, text=True, capture_output=True, timeout=timeout)
    if check and process.returncode:
        raise RuntimeError(process.stderr or process.stdout)
    return process


def office(*args, check=True, timeout=300):
    if RUNNER is None:
        raise RuntimeError("thesis adapter has not been bound to the pipeline runner")
    return RUNNER.office(*args, check=check, timeout=timeout)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _disable_automatic_field_refresh(path: Path) -> None:
    """Remove Word's open-time field refresh trigger from a DOCX package.

    The pipeline refreshes only TOC/SEQ/REF/PAGEREF/PAGE fields explicitly in
    Microsoft Word. Keeping w:updateFields would trigger a modal prompt before
    automation can skip Zotero and legacy form fields.
    """
    temporary = path.with_name(f".{path.name}.{uuid.uuid4().hex}.tmp")
    changed = False
    with zipfile.ZipFile(path, "r") as source_zip, zipfile.ZipFile(
        temporary, "w", compression=zipfile.ZIP_DEFLATED
    ) as target_zip:
        for info in source_zip.infolist():
            payload = source_zip.read(info.filename)
            if info.filename == "word/settings.xml":
                root = etree.fromstring(payload)
                for node in root.findall(qn("w:updateFields")):
                    root.remove(node)
                    changed = True
                payload = etree.tostring(
                    root, xml_declaration=True, encoding="UTF-8", standalone=True
                )
            target_zip.writestr(info, payload)
    if changed:
        shutil.copystat(path, temporary)
        temporary.replace(path)
    else:
        temporary.unlink(missing_ok=True)


@contextmanager
def _word_process_lock(lock_path: Path, timeout_seconds: float):
    """Cross-process file lock for Microsoft Word automation.

    Uses fcntl.flock on macOS/Linux or falls back to polling-based lock file
    when fcntl is unavailable. Includes owner PID, path, and wait diagnostics
    in timeout errors.
    """
    if isinstance(timeout_seconds, bool) or timeout_seconds <= 0:
        raise ValueError("timeout_seconds must be positive")
    lock_path.parent.mkdir(parents=True, exist_ok=True)
    owner_pid = os.getpid()
    acquired = False
    start_time = time.monotonic()

    if HAS_FCNTL:
        # fcntl-based lock for POSIX systems
        lock_file = None
        try:
            # The path is deliberately persistent. Opening with "w" before
            # flock would let a waiter truncate the current owner's metadata;
            # unlinking after unlock would let contenders lock different
            # inodes. Both break process-wide mutual exclusion.
            lock_file = lock_path.open("a+", encoding="utf-8")
            deadline = start_time + timeout_seconds
            while True:
                try:
                    fcntl.flock(lock_file.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                    acquired = True
                    break
                except (IOError, OSError):
                    if time.monotonic() >= deadline:
                        break
                    time.sleep(0.1)

            if not acquired:
                waited = time.monotonic() - start_time
                existing_owner = "unknown"
                try:
                    lock_file.seek(0)
                    existing_owner = lock_file.read().strip() or "unknown"
                except Exception:
                    pass
                raise TimeoutError(
                    f"Failed to acquire Word lock after {waited:.1f}s. "
                    f"Lock path: {lock_path}, existing owner: {existing_owner}, "
                    f"current PID: {owner_pid}"
                )

            lock_file.seek(0)
            lock_file.truncate()
            lock_file.write(json.dumps({
                "pid": owner_pid,
                "path": str(lock_path),
                "acquired_at_epoch": time.time(),
            }, sort_keys=True))
            lock_file.flush()
            os.fsync(lock_file.fileno())
            yield
        finally:
            if lock_file is not None:
                if acquired:
                    try:
                        fcntl.flock(lock_file.fileno(), fcntl.LOCK_UN)
                    except Exception:
                        pass
                lock_file.close()
    else:
        # Fallback polling-based lock for Windows
        try:
            deadline = start_time + timeout_seconds
            while True:
                try:
                    with lock_path.open("x", encoding="utf-8") as handle:
                        handle.write(json.dumps({
                            "pid": owner_pid,
                            "path": str(lock_path),
                            "acquired_at_epoch": time.time(),
                        }, sort_keys=True))
                    acquired = True
                    break
                except FileExistsError:
                    if time.monotonic() >= deadline:
                        break
                    time.sleep(0.1)

            if not acquired:
                waited = time.monotonic() - start_time
                existing_owner = "unknown"
                try:
                    with open(lock_path, "r") as f:
                        existing_owner = f.read().strip()
                except Exception:
                    pass
                raise TimeoutError(
                    f"Failed to acquire Word lock after {waited:.1f}s. "
                    f"Lock path: {lock_path}, existing owner: {existing_owner}, "
                    f"current PID: {owner_pid}"
                )

            yield
        finally:
            if acquired:
                lock_path.unlink(missing_ok=True)


def prepare_word_staging(path: Path, *, purpose: str) -> None:
    _disable_automatic_field_refresh(path)


def _profile_path(value: str) -> Path:
    candidate = Path(value)
    if candidate.is_file():
        return candidate.resolve()
    bundled = PROFILE_DIR / f"{value}.json"
    if bundled.is_file():
        return bundled
    raise FileNotFoundError(f"Unknown thesis profile: {value}")


def load_profile(job: dict) -> tuple[Path, dict]:
    path = _profile_path(job["profile"])
    return path, json.loads(path.read_text(encoding="utf-8"))


def configure(job: dict, run_dir: Path) -> dict[str, str]:
    artifacts = run_dir / "artifacts"
    evidence = run_dir / "evidence"
    artifacts.mkdir(parents=True, exist_ok=True)
    evidence.mkdir(parents=True, exist_ok=True)
    stem = Path(job["source"]).stem
    profile_id = Path(str(job["profile"])).stem
    delivery_date = job.get("output", {}).get("delivery_date")
    if delivery_date and profile_id == "ruc-doctoral-2026":
        clean_name = f"{stem}_RUC格式化清洁稿_{delivery_date}.docx"
        tracked_name = f"{stem}_RUC格式修订稿_{delivery_date}.docx"
    elif delivery_date:
        clean_name = f"{stem}_{profile_id}_formatted_{delivery_date}.docx"
        tracked_name = f"{stem}_{profile_id}_format_tracking_{delivery_date}.docx"
    else:
        # Preserve path compatibility for jobs created before an explicit,
        # stable delivery date was added to schema 2.0.
        clean_name = f"{stem}_{profile_id}_formatted.docx"
        tracked_name = f"{stem}_{profile_id}_format_tracking.docx"
    paths = {
        "prepared": str(artifacts / f"{stem}_layout_candidate.docx"),
        "metadata": str(evidence / "thesis_layout_metadata.json"),
        "layout_pdf": str(evidence / "thesis_layout_candidate.pdf"),
        "accepted": str(artifacts / clean_name),
        "tracked": str(artifacts / tracked_name),
        "batch": str(evidence / "thesis_officecli_batch.json"),
        "command_count": str(evidence / "thesis_command_count.txt"),
        "accepted_pdf": str(artifacts / clean_name.replace(".docx", ".pdf")),
        "format_diff_json": str(run_dir / "audit" / "format_diff.json"),
        "format_diff_csv": str(run_dir / "audit" / "format_diff.csv"),
        "style_proposal": str(run_dir / "audit" / "style_mapping_proposal.json"),
        "semantic_before": str(evidence / "semantic_before.json"),
        "semantic_after": str(evidence / "semantic_after.json"),
        "refresh_log": str(evidence / "word_field_refresh.json"),
        "adapter_qa_json": str(evidence / "thesis_format_qa.json"),
        "qa_md": str(run_dir / "thesis_format_QA.md"),
    }
    return paths


def _office_results(payload: dict) -> list[dict]:
    return payload.get("data", {}).get("results", []) if payload.get("success") else []


def _canonical(value: Any) -> str:
    if isinstance(value, bool):
        return str(value).lower()
    if value is None:
        return ""
    return str(value).strip().lower()


def _format_equivalent(prop: str, current: Any, target: Any) -> bool:
    current_value = _canonical(current)
    target_value = _canonical(target)
    if prop in {"docDefaults.spaceBefore", "docDefaults.spaceAfter"}:
        zero_values = {"", "0", "0pt", "0.0", "0.0pt"}
        return current_value in zero_values and target_value in zero_values
    return current_value == target_value


STYLE_HINTS = {
    "body": ("正文", "body", "normal"),
    "chapter": ("章标题", "chapter", "标题 1", "heading 1"),
    "heading1": ("一级", "二级标题", "heading 2", "标题 2"),
    "heading2": ("二级", "三级标题", "heading 3", "标题 3"),
    "heading3": ("三级", "四级标题", "heading 4", "标题 4"),
    "caption": ("题注", "caption", "图题", "表题"),
    "footnote": ("脚注", "footnote"),
    "bibliography": ("参考文献", "bibliography"),
}


def _style_proposals(style_rows: list[dict]) -> dict[str, Any]:
    proposals = {}
    for role, hints in STYLE_HINTS.items():
        candidates = []
        for row in style_rows:
            haystack = f"{row.get('path', '')} {row.get('text', '')} {row.get('format', {}).get('name', '')}".lower()
            if any(hint.lower() in haystack for hint in hints):
                candidates.append(row.get("path"))
        proposals[role] = {
            "candidates": sorted(set(value for value in candidates if value)),
            "unambiguous": len(set(candidates)) == 1,
            "applied_automatically": False,
        }
    return proposals


def augment_audit(job: dict, source: Path, output_dir: Path, results: dict) -> dict:
    profile_path, profile = load_profile(job)
    document_rows = _office_results(results.get("document", {}))
    document_format = document_rows[0].get("format", {}) if document_rows else {}
    style_rows = _office_results(results.get("styles", {}))
    proposals = _style_proposals(style_rows)
    style_paths = {row.get("path") for row in style_rows}
    configured = job.get("style_map", {})
    rows: list[dict[str, Any]] = []
    for prop, target in profile.get("document_props", {}).items():
        current = document_format.get(prop)
        rows.append({
            "rule_id": "RUC-PAGE-001",
            "area": "document",
            "target": prop,
            "current": current,
            "required": target,
            "status": "pass" if _format_equivalent(prop, current, target) else "change",
            "action": "officecli set /",
            "confidence": "high",
        })
    for role in profile.get("style_roles", {}):
        selector = configured.get(role, "")
        exists = selector in style_paths if selector else False
        rows.append({
            "rule_id": "RUC-STYLE-001",
            "area": "style_mapping",
            "target": role,
            "current": selector,
            "required": "explicit existing style path",
            "status": "pass" if exists else "unresolved",
            "action": "configure style_map",
            "confidence": "high" if exists else "requires_review",
        })
    document = Document(source)
    body_text = [paragraph.text.strip() for paragraph in document.paragraphs if paragraph.text.strip()]
    expected_order = ["中文摘要", "英文摘要", "目录", "参考文献", "致谢"]
    found = {label: next((index for index, text in enumerate(body_text) if label in text), None) for label in expected_order}
    ordered_positions = [value for value in found.values() if value is not None]
    structure_ok = ordered_positions == sorted(ordered_positions)

    # For full_thesis scope, missing major blocks should be reported as "review"
    scope = job.get("scope", "chapter")
    missing_blocks = [label for label, pos in found.items() if pos is None]
    if scope == "full_thesis" and missing_blocks:
        structure_status = "review"
        structure_current = json.dumps({"found": found, "missing": missing_blocks}, ensure_ascii=False)
    else:
        structure_status = "pass" if structure_ok else "review"
        structure_current = json.dumps(found, ensure_ascii=False)

    rows.append({
        "rule_id": "RUC-STRUCT-001",
        "area": "structure",
        "target": "front/body/back matter order",
        "current": structure_current,
        "required": "validate supplied major-block anchors",
        "status": structure_status,
        "action": "validate_only" if job.get("structure_policy") == "validate_only" else "explicit anchors",
        "confidence": "diagnostic",
    })
    diff_json = output_dir / "format_diff.json"
    diff_csv = output_dir / "format_diff.csv"
    proposal_json = output_dir / "style_mapping_proposal.json"
    diff_json.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    with diff_csv.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)
    proposal_json.write_text(json.dumps(proposals, ensure_ascii=False, indent=2), encoding="utf-8")
    return {
        "profile": str(profile_path),
        "profile_sha256": sha256(profile_path),
        "format_diff": str(diff_json),
        "style_mapping_proposal": str(proposal_json),
        "unresolved_style_roles": [row["target"] for row in rows if row["area"] == "style_mapping" and row["status"] != "pass"],
        "structure_policy": job.get("structure_policy", "validate_only"),
        "structure_reordered": False,
    }


def _xml_semantic_snapshot(path: Path) -> dict[str, Any]:
    ns = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    }
    with zipfile.ZipFile(path) as package:
        names = set(package.namelist())
        root = etree.fromstring(package.read("word/document.xml"))
        for element in root.xpath(".//w:sectPr", namespaces=ns):
            parent = element.getparent()
            if parent is not None:
                parent.remove(element)
        # Field caches (TOC, page numbers, cross-references, Zotero display text)
        # legitimately change during Word refresh. A TOC field can span many
        # paragraphs, so track field depth in document order.
        for field in root.xpath(".//w:fldSimple", namespaces=ns):
            parent = field.getparent()
            if parent is not None:
                parent.remove(field)
        field_depth = 0
        for run_element in list(root.xpath(".//w:r", namespaces=ns)):
            field_chars = run_element.xpath(".//w:fldChar", namespaces=ns)
            kinds = [node.get(qn("w:fldCharType")) for node in field_chars]
            remove = field_depth > 0 or "begin" in kinds
            field_depth += kinds.count("begin")
            field_depth -= kinds.count("end")
            if remove or "end" in kinds:
                parent = run_element.getparent()
                if parent is not None:
                    parent.remove(run_element)
        # Word may split an unchanged sentence into additional runs while
        # refreshing fields. Join runs within each paragraph before hashing so
        # the invariant measures semantic text, not run segmentation.
        paragraph_texts = []
        for paragraph in root.xpath(".//w:p", namespaces=ns):
            text = "".join(paragraph.xpath(".//w:t/text()", namespaces=ns))
            if text:
                paragraph_texts.append(text)
        paragraph_text = "\n".join(paragraph_texts)

        # Normalize footnote text independently of run segmentation
        footnote_text = ""
        if "word/footnotes.xml" in names:
            footnotes = etree.fromstring(package.read("word/footnotes.xml"))
            footnote_paragraphs = []
            for fn_para in footnotes.xpath(".//w:p", namespaces=ns):
                fn_text = "".join(fn_para.xpath(".//w:t/text()", namespaces=ns))
                if fn_text:
                    footnote_paragraphs.append(fn_text)
            footnote_text = "\n".join(footnote_paragraphs)

        # Hash both the package payload multiset and each drawing's ordered
        # relationship to its payload. The latter detects relationship swaps
        # while remaining invariant to harmless relationship-id renumbering.
        media_hashes = []
        for member in sorted(names):
            if member.startswith("word/media/"):
                media_bytes = package.read(member)
                media_hashes.append(hashlib.sha256(media_bytes).hexdigest())
        media_payload_sha256 = hashlib.sha256(
            json.dumps(media_hashes, separators=(",", ":")).encode("utf-8")
        ).hexdigest()

        media_usage: list[tuple[str, int, str]] = []
        equation_usage: list[tuple[str, int, str]] = []
        for member in sorted(names):
            if not member.startswith("word/") or not member.endswith(".xml"):
                continue
            try:
                part = etree.fromstring(package.read(member))
            except etree.XMLSyntaxError:
                continue

            part_path = PurePosixPath(member)
            rels_name = str(part_path.parent / "_rels" / f"{part_path.name}.rels")
            relationships: dict[str, str] = {}
            if rels_name in names:
                try:
                    rels_root = etree.fromstring(package.read(rels_name))
                    for relationship in rels_root.xpath("./pr:Relationship", namespaces=ns):
                        if relationship.get("TargetMode") == "External":
                            continue
                        rel_id = relationship.get("Id")
                        target = relationship.get("Target")
                        if rel_id and target:
                            if target.startswith("/"):
                                resolved = target.lstrip("/")
                            else:
                                resolved = posixpath.normpath(
                                    posixpath.join(str(part_path.parent), target)
                                )
                            relationships[rel_id] = resolved
                except etree.XMLSyntaxError:
                    relationships = {}
            for index, blip in enumerate(part.xpath(".//a:blip[@r:embed]", namespaces=ns)):
                rel_id = blip.get(f"{{{ns['r']}}}embed")
                target = relationships.get(rel_id or "", "")
                payload_hash = (
                    hashlib.sha256(package.read(target)).hexdigest()
                    if target in names
                    else f"unresolved:{rel_id or ''}"
                )
                media_usage.append((member, index, payload_hash))

            for index, omath in enumerate(part.xpath(".//m:oMath", namespaces=ns)):
                canonical = etree.tostring(
                    omath, method="c14n", exclusive=False, with_comments=False
                )
                equation_usage.append(
                    (member, index, hashlib.sha256(canonical).hexdigest())
                )

        media_usage_sha256 = hashlib.sha256(
            json.dumps(media_usage, separators=(",", ":")).encode("utf-8")
        ).hexdigest()
        equation_xml_sha256 = hashlib.sha256(
            json.dumps(equation_usage, separators=(",", ":")).encode("utf-8")
        ).hexdigest()

        instructions = []
        for member in names:
            if member.startswith("word/") and member.endswith(".xml"):
                try:
                    part = etree.fromstring(package.read(member))
                except etree.XMLSyntaxError:
                    continue
                for paragraph in part.xpath(".//w:p", namespaces=ns):
                    instruction = "".join(paragraph.xpath(".//w:instrText/text()", namespaces=ns))
                    if instruction:
                        instructions.append(instruction)
        zotero = sorted(text for text in instructions if "ADDIN ZOTERO_ITEM" in text or "ADDIN ZOTERO_BIBL" in text)
        return {
            "body_text_sha256": hashlib.sha256(paragraph_text.encode("utf-8")).hexdigest(),
            "footnote_text_sha256": hashlib.sha256(footnote_text.encode("utf-8")).hexdigest(),
            "media_payload_sha256": media_payload_sha256,
            "media_usage_sha256": media_usage_sha256,
            "equation_xml_sha256": equation_xml_sha256,
            "paragraph_text_length": len(paragraph_text),
            "tables": len(root.xpath(".//w:tbl", namespaces=ns)),
            "equations": len(root.xpath(".//m:oMath", namespaces=ns)),
            "drawings": len(root.xpath(".//w:drawing", namespaces=ns)),
            "inline_drawings": len(root.xpath(".//wp:inline", namespaces=ns)),
            "floating_drawings": len(root.xpath(".//wp:anchor", namespaces=ns)),
            "zotero_fields": len(zotero),
            "zotero_fields_sha256": hashlib.sha256("\n".join(zotero).encode("utf-8")).hexdigest(),
        }


def _ensure_setting(settings, name: str) -> None:
    if settings.find(qn(f"w:{name}")) is None:
        settings.append(OxmlElement(f"w:{name}"))


def _set_rfonts(run, east_asia: str, latin: str) -> None:
    rpr = run._element.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.insert(0, fonts)
    fonts.set(qn("w:eastAsia"), east_asia)
    fonts.set(qn("w:ascii"), latin)
    fonts.set(qn("w:hAnsi"), latin)


def _replace_story_text(story, text: str, *, font_ea: str, font_latin: str, size: float, alignment) -> None:
    for paragraph in list(story.paragraphs):
        paragraph._element.getparent().remove(paragraph._element)
    paragraph = story.add_paragraph()
    paragraph.alignment = alignment
    run = paragraph.add_run(text)
    run.font.name = font_latin
    run.font.size = Pt(size)
    _set_rfonts(run, font_ea, font_latin)


def _replace_footer_with_page(story, *, font_latin: str, size: float, alignment) -> None:
    for paragraph in list(story.paragraphs):
        paragraph._element.getparent().remove(paragraph._element)
    paragraph = story.add_paragraph()
    paragraph.alignment = alignment
    run = paragraph.add_run()
    _set_rfonts(run, font_latin, font_latin)
    run.font.size = Pt(size)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instruction.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    result = OxmlElement("w:t")
    result.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._element.extend([begin, instruction, separate, result, end])


def _set_page_number(section, fmt: str, start: int | None = None) -> None:
    sect_pr = section._sectPr
    current = sect_pr.find(qn("w:pgNumType"))
    if current is None:
        current = OxmlElement("w:pgNumType")
        sect_pr.append(current)
    current.set(qn("w:fmt"), fmt)
    if start is not None:
        current.set(qn("w:start"), str(start))


def _set_border(parent, edge: str, *, value: str, size: str = "0") -> None:
    borders = parent.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        parent.append(borders)
    node = borders.find(qn(f"w:{edge}"))
    if node is None:
        node = OxmlElement(f"w:{edge}")
        borders.append(node)
    node.set(qn("w:val"), value)
    if value != "nil":
        node.set(qn("w:sz"), size)
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), "auto")


def _format_three_line_table(table, profile: dict) -> None:
    tbl_pr = table._tbl.tblPr
    for edge in ("left", "right", "insideH", "insideV"):
        _set_border(tbl_pr, edge, value="nil")
    _set_border(tbl_pr, "top", value="single", size="12")
    _set_border(tbl_pr, "bottom", value="single", size="12")
    if table.rows:
        for cell in table.rows[0].cells:
            tc_pr = cell._tc.get_or_add_tcPr()
            borders = tc_pr.find(qn("w:tcBorders"))
            if borders is None:
                borders = OxmlElement("w:tcBorders")
                tc_pr.append(borders)
            bottom = OxmlElement("w:bottom")
            bottom.set(qn("w:val"), "single")
            bottom.set(qn("w:sz"), "6")
            bottom.set(qn("w:space"), "0")
            bottom.set(qn("w:color"), "auto")
            borders.append(bottom)
    table_spec = profile["table"]
    for row in table.rows:
        tr_pr = row._tr.get_or_add_trPr()
        if tr_pr.find(qn("w:cantSplit")) is None:
            tr_pr.append(OxmlElement("w:cantSplit"))
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(table_spec["size_pt"])
                    _set_rfonts(run, table_spec["font_east_asia"], table_spec["font_latin"])


def _apply_structural_formatting(path: Path, job: dict, profile: dict) -> list[str]:
    document = Document(path)
    warnings = []
    _ensure_setting(document.settings.element, "mirrorMargins")
    document.settings.odd_and_even_pages_header_footer = True

    # Use the canonical profile property names consumed by OfficeCLI.
    page_props = profile.get("document_props", {})

    def _parse_length(value: Any):
        """Parse an explicit OOXML-facing length; bare numbers mean cm."""
        if isinstance(value, bool):
            raise ValueError(f"Invalid profile length: {value!r}")
        match = re.fullmatch(
            r"\s*([0-9]+(?:\.[0-9]+)?)\s*(cm|mm|in|pt)?\s*",
            str(value).lower(),
        )
        if not match:
            raise ValueError(f"Invalid profile length: {value!r}")
        number = float(match.group(1))
        unit = match.group(2) or "cm"
        return {"cm": Cm, "mm": Mm, "in": Inches, "pt": Pt}[unit](number)

    page_width = _parse_length(page_props.get("pageWidth", "21cm"))
    page_height = _parse_length(page_props.get("pageHeight", "29.7cm"))
    margin_top = _parse_length(page_props.get("marginTop", "4.5cm"))
    margin_bottom = _parse_length(page_props.get("marginBottom", "4cm"))
    margin_left = _parse_length(page_props.get("marginLeft", "3.5cm"))
    margin_right = _parse_length(page_props.get("marginRight", "3cm"))
    margin_header = _parse_length(page_props.get("marginHeader", "2.5cm"))
    margin_footer = _parse_length(page_props.get("marginFooter", "2.5cm"))

    for section in document.sections:
        landscape = section.page_width and section.page_height and section.page_width > section.page_height
        if landscape:
            section.page_width = page_height
            section.page_height = page_width
        else:
            section.page_width = page_width
            section.page_height = page_height
        section.top_margin = margin_top
        section.bottom_margin = margin_bottom
        section.left_margin = margin_left
        section.right_margin = margin_right
        section.header_distance = margin_header
        section.footer_distance = margin_footer

    title = job.get("thesis_title", "").strip()
    if not title:
        warnings.append("thesis_title_missing_headers_preserved")
    else:
        header = profile["header"]
        footer = profile["footer"]
        for section in document.sections:
            section.header.is_linked_to_previous = False
            section.even_page_header.is_linked_to_previous = False
            section.footer.is_linked_to_previous = False
            section.even_page_footer.is_linked_to_previous = False
            _replace_story_text(section.header, title, font_ea=header["font_east_asia"], font_latin=header["font_latin"], size=header["size_pt"], alignment=WD_ALIGN_PARAGRAPH.CENTER)
            _replace_story_text(section.even_page_header, title, font_ea=header["font_east_asia"], font_latin=header["font_latin"], size=header["size_pt"], alignment=WD_ALIGN_PARAGRAPH.CENTER)
            _replace_footer_with_page(section.footer, font_latin=footer["font_latin"], size=footer["size_pt"], alignment=WD_ALIGN_PARAGRAPH.RIGHT)
            _replace_footer_with_page(section.even_page_footer, font_latin=footer["font_latin"], size=footer["size_pt"], alignment=WD_ALIGN_PARAGRAPH.LEFT)
    if job.get("scope") == "full_thesis":
        section_map = job.get("section_map", {})
        front = section_map.get("front_matter_section")
        body = section_map.get("body_section")
        if front is not None and body is not None:
            if not (1 <= int(front) <= len(document.sections) and 1 <= int(body) <= len(document.sections)):
                raise ValueError("section_map uses 1-based section numbers outside the document")
            _set_page_number(document.sections[int(front) - 1], profile["page_numbering"]["front_matter"], 1)
            _set_page_number(document.sections[int(body) - 1], profile["page_numbering"]["body"], profile["page_numbering"]["body_start"])
        else:
            warnings.append("page_numbering_preserved_missing_explicit_section_map")
    table_policy = job.get("table_policy", "report_only")
    if table_policy == "all_body_tables":
        targets = range(len(document.tables))
    elif table_policy == "explicit":
        targets = [int(value) - 1 for value in job.get("table_indices", [])]
    else:
        targets = []
    for index in targets:
        if not 0 <= index < len(document.tables):
            raise ValueError(f"table index outside document: {index + 1}")
        _format_three_line_table(document.tables[index], profile)
    document.save(path)
    return warnings


def _field_operations(job: dict) -> list[dict]:
    policy = job.get("field_policy", "preserve")
    operations = job.get("field_operations", [])
    if policy == "preserve":
        if operations:
            raise ValueError("field_operations are forbidden when field_policy=preserve")
        return []
    if not operations:
        raise ValueError(f"field_policy={policy} requires explicit field_operations")
    allowed_types = {"field", "toc"}
    validated = []
    for operation in operations:
        if operation.get("command") != "add" or operation.get("type") not in allowed_types or not operation.get("parent"):
            raise ValueError("field_operations may only add field/toc elements at explicit parents")
        validated.append(operation)
    return validated


def _require_style_map(job: dict, profile: dict, source: Path) -> None:
    mappings = job.get("style_map", {})
    if not mappings:
        raise ValueError("No style_map configured. Run audit and explicitly map thesis styles before layout.")

    document = Document(source)
    available_style_paths = {
        f"/styles/{style.style_id}"
        for style in document.styles
        if style.style_id
    }

    # Validate each configured selector exists in source
    for role, selector in mappings.items():
        if selector and selector not in available_style_paths:
            raise ValueError(
                f"Style selector '{selector}' for role '{role}' does not exist in source document. "
                f"Available styles: {sorted(available_style_paths)}"
            )

    required = set(job.get("required_style_roles", ["body"]))
    if job.get("scope") == "full_thesis":
        required.update(job.get("full_thesis_required_style_roles", ["chapter", "heading1", "heading2", "caption"]))
    missing = sorted(role for role in required if role not in mappings)
    if missing:
        raise ValueError(f"Missing explicit style_map roles: {missing}")
    invalid_roles = sorted(role for role in mappings if role not in profile.get("style_roles", {}))
    if invalid_roles:
        raise ValueError(f"Unknown profile style roles: {invalid_roles}")


def build_layout(job: dict, paths: dict[str, str]) -> dict:
    source = Path(job["source"])
    prepared = Path(paths["prepared"])
    profile_path, profile = load_profile(job)
    _require_style_map(job, profile, source)
    before = _xml_semantic_snapshot(source)
    Path(paths["semantic_before"]).write_text(json.dumps(before, ensure_ascii=False, indent=2), encoding="utf-8")
    shutil.copy2(source, prepared)
    commands = [{"command": "set", "path": "/", "props": profile["document_props"]}]
    for role, selector in job.get("style_map", {}).items():
        commands.append({"command": "set", "path": selector, "props": profile["style_roles"][role]})
    commands.extend(_field_operations(job))
    batch = Path(paths["batch"])
    batch.write_text(json.dumps(commands, ensure_ascii=False, indent=2), encoding="utf-8")
    result = office("batch", str(prepared), "--input", str(batch), "--json", timeout=900)
    office("close", str(prepared), "--json", check=False)
    warnings = _apply_structural_formatting(prepared, job, profile)
    _disable_automatic_field_refresh(prepared)
    after = _xml_semantic_snapshot(prepared)
    if before != after:
        changed = [key for key in before if before[key] != after.get(key)]
        allowed = {"body_text_sha256", "paragraph_text_length"} if job.get("field_policy") != "preserve" else set()
        unexpected = [key for key in changed if key not in allowed]
        if unexpected:
            raise RuntimeError(f"Thesis formatting changed semantic invariants: {unexpected}")
    Path(paths["command_count"]).write_text(str(len(commands)), encoding="utf-8")
    return {
        "profile": str(profile_path),
        "profile_sha256": sha256(profile_path),
        "source_sha256": sha256(source),
        "prepared_sha256": sha256(prepared),
        "officecli_commands": len(commands),
        "officecli_stdout": result.stdout.strip(),
        "warnings": warnings,
        "semantic_before": before,
        "semantic_after_layout": after,
        "structure_reordered": False,
    }


def restore_layout_sidecars(metadata: dict[str, Any], paths: dict[str, str]) -> None:
    """Restore run-local evidence omitted from the shared binary layout cache."""
    semantic_path = Path(paths["semantic_before"])
    semantic_path.parent.mkdir(parents=True, exist_ok=True)
    semantic_path.write_text(
        json.dumps(metadata["semantic_before"], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    Path(paths["command_count"]).write_text(
        str(metadata.get("officecli_commands", 0)), encoding="utf-8"
    )


def _word_lock_path() -> Path:
    return Path(tempfile.gettempdir()) / "officecli-word-revision" / "word-automation.lock"


def _word_compare_script() -> str:
    return r'''
on run argv
  set originalName to item 1 of argv
  set revisedName to item 2 of argv
  set outputPath to item 3 of argv
  set revisionAuthor to item 4 of argv
  with timeout of 900 seconds
  tell application "Microsoft Word"
    set previousAlerts to display alerts
    set originalDoc to missing value
    set revisedDoc to missing value
    set comparedDoc to missing value
    try
      set display alerts to alerts none
      repeat 240 times
        try
          set originalDoc to document originalName
          set revisedDoc to document revisedName
          exit repeat
        on error
          delay 0.25
        end try
      end repeat
      if originalDoc is missing value or revisedDoc is missing value then error "Word comparison inputs did not open"
      set revisedPath to full name of revisedDoc
      close revisedDoc saving no
      set revisedDoc to missing value
      compare originalDoc path revisedPath author name revisionAuthor target compare target new detect format changes true ignore all comparison warnings true add to recent files false
      set comparedDoc to active document
      save as comparedDoc file name POSIX file outputPath file format format document default add to recent files false
      set revisionCount to count of revisions of comparedDoc
      close comparedDoc saving no
      set comparedDoc to missing value
      close originalDoc saving no
      set originalDoc to missing value
      set display alerts to previousAlerts
    on error errMsg number errNum
      try
        if comparedDoc is not missing value then close comparedDoc saving no
      end try
      try
        if revisedDoc is not missing value then close revisedDoc saving no
      end try
      try
        if originalDoc is not missing value then close originalDoc saving no
      end try
      set display alerts to previousAlerts
      error errMsg number errNum
    end try
  end tell
  end timeout
  return "revisions=" & revisionCount
end run
'''


def _word_compare_formatting_unlocked(original: Path, revised: Path, output: Path, author: str) -> dict[str, Any]:
    run(["open", "-a", "Microsoft Word", str(original)], timeout=60)
    run(["open", "-a", "Microsoft Word", str(revised)], timeout=60)
    time.sleep(3)
    script = _word_compare_script()
    with tempfile.NamedTemporaryFile("w", suffix=".applescript", encoding="utf-8", delete=False) as handle:
        handle.write(script)
        script_path = Path(handle.name)
    try:
        process = run(
            ["osascript", str(script_path), original.name, revised.name, str(output), author],
            timeout=900,
        )
    finally:
        script_path.unlink(missing_ok=True)
    match = re.search(r"revisions=(\d+)", process.stdout)
    return {
        "stdout": process.stdout.strip(),
        "revisions": int(match.group(1)) if match else None,
    }


def _word_compare_formatting(original: Path, revised: Path, output: Path, author: str) -> dict[str, Any]:
    with _word_process_lock(_word_lock_path(), WORD_LOCK_TIMEOUT_SECONDS):
        return _word_compare_formatting_unlocked(original, revised, output, author)


def build_tracked(job: dict, paths: dict[str, str], metadata: dict) -> None:
    if not job.get("output", {}).get("tracked_formatting", False):
        return
    original = Path(job["source"])
    revised = Path(paths["prepared"])
    tracked = Path(paths["tracked"])
    tracked.parent.mkdir(parents=True, exist_ok=True)
    staging_dir = Path(job.get("word_pdf_staging_dir", tracked.parent)).expanduser().resolve()
    staging_dir.mkdir(parents=True, exist_ok=True)
    safe_job_id = re.sub(r"[^A-Za-z0-9_.-]+", "-", job["job_id"])
    nonce = uuid.uuid4().hex[:10]
    staged_original = staging_dir / f"wordrev_{safe_job_id}_compare_original_{nonce}.docx"
    staged_revised = staging_dir / f"wordrev_{safe_job_id}_compare_revised_{nonce}.docx"
    staged_output = staging_dir / f"wordrev_{safe_job_id}_compare_tracked_{nonce}.docx"
    for path in (staged_original, staged_revised, staged_output):
        path.unlink(missing_ok=True)
    try:
        shutil.copy2(original, staged_original)
        shutil.copy2(revised, staged_revised)
        _disable_automatic_field_refresh(staged_original)
        _disable_automatic_field_refresh(staged_revised)
        comparison = _word_compare_formatting(
            staged_original,
            staged_revised,
            staged_output,
            job.get("revision_author", "Codex"),
        )
        if not staged_output.is_file():
            raise RuntimeError("Microsoft Word did not produce the requested format-tracking document")
        if comparison.get("revisions") == 0:
            raise RuntimeError("Microsoft Word comparison produced no formatting revisions")
        shutil.copy2(staged_output, tracked)
        metadata.setdefault("tracked_formatting", {}).update(comparison)
        Path(paths["metadata"]).write_text(
            json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    finally:
        for path in (staged_original, staged_revised, staged_output):
            path.unlink(missing_ok=True)


def _word_refresh_script() -> str:
    return r'''
on run argv
  set inputName to item 1 of argv
  with timeout of 900 seconds
  tell application "Microsoft Word"
    set previousAlerts to display alerts
    set docRef to missing value
    try
      set display alerts to alerts none
      repeat 240 times
        try
          set docRef to document inputName
          exit repeat
        on error
          delay 0.25
        end try
      end repeat
      if docRef is missing value then error "Word refresh input did not open"
      set mainText to text object of docRef
      set fieldCount to 0
      set updatedCount to 0
      set skippedCount to 0
      set errorCount to 0
      set fieldIndex to 1
      repeat
        try
          set fieldRef to field fieldIndex of mainText
        on error
          exit repeat
        end try
        set fieldCount to fieldCount + 1
        set codeText to ""
        try
          set codeText to content of field code of fieldRef
        end try
        if codeText contains "ADDIN ZOTERO" or codeText contains "FORMTEXT" or codeText contains "FORMCHECKBOX" or codeText contains "FORMDROPDOWN" then
          set skippedCount to skippedCount + 1
        else if codeText starts with " TOC " or codeText starts with "TOC " or codeText starts with " SEQ " or codeText starts with "SEQ " or codeText starts with " REF " or codeText starts with "REF " or codeText starts with " PAGEREF " or codeText starts with "PAGEREF " or codeText starts with " PAGE " or codeText starts with "PAGE " or codeText starts with " NUMPAGES " or codeText starts with "NUMPAGES " or codeText starts with " STYLEREF " or codeText starts with "STYLEREF " then
          try
            update field fieldRef
            set updatedCount to updatedCount + 1
          on error
            set errorCount to errorCount + 1
          end try
        else
          set skippedCount to skippedCount + 1
        end if
        set fieldIndex to fieldIndex + 1
        if fieldIndex > 20000 then exit repeat
      end repeat
      set tocIndex to 1
      repeat
        try
          set tocRef to table of contents tocIndex of docRef
        on error
          exit repeat
        end try
        try
          update tocRef
          update page numbers tocRef
        end try
        set tocIndex to tocIndex + 1
      end repeat
      set tofIndex to 1
      repeat
        try
          set tofRef to table of figures tofIndex of docRef
        on error
          exit repeat
        end try
        try
          update tofRef
          update page numbers tofRef
        end try
        set tofIndex to tofIndex + 1
      end repeat
      save docRef
      close docRef saving no
      set docRef to missing value
      set display alerts to previousAlerts
    on error errMsg number errNum
      try
        if docRef is not missing value then close docRef saving no
      end try
      set display alerts to previousAlerts
      error errMsg number errNum
    end try
  end tell
  end timeout
  return "fields=" & fieldCount & ",updated=" & updatedCount & ",skipped=" & skippedCount & ",errors=" & errorCount
end run
'''


def _word_refresh_selected_unlocked(path: Path) -> dict[str, Any]:
    run(["open", "-a", "Microsoft Word", str(path)], timeout=60)
    time.sleep(3)
    script = _word_refresh_script()
    with tempfile.NamedTemporaryFile("w", suffix=".applescript", encoding="utf-8", delete=False) as handle:
        handle.write(script)
        script_path = Path(handle.name)
    try:
        process = run(["osascript", str(script_path), path.name], timeout=900)
    finally:
        script_path.unlink(missing_ok=True)
    values = {}
    for pair in process.stdout.strip().split(","):
        if "=" in pair:
            key, value = pair.split("=", 1)
            values[key] = int(value) if value.isdigit() else value
    return {"stdout": process.stdout.strip(), **values}


def _word_refresh_selected(path: Path) -> dict[str, Any]:
    with _word_process_lock(_word_lock_path(), WORD_LOCK_TIMEOUT_SECONDS):
        return _word_refresh_selected_unlocked(path)


def postprocess(job: dict, paths: dict[str, str]) -> None:
    prepared = Path(paths["prepared"])
    accepted = Path(paths["accepted"])
    shutil.copy2(prepared, accepted)
    refresh = {"execution_skipped": True, "reason": "chapter scope or no target fields"}
    if job.get("scope") == "full_thesis" or job.get("field_policy") != "preserve":
        staging_value = job.get("word_pdf_staging_dir")
        if staging_value:
            staging_dir = Path(staging_value).expanduser().resolve()
            staging_dir.mkdir(parents=True, exist_ok=True)
            safe_job_id = re.sub(r"[^A-Za-z0-9_.-]+", "-", job["job_id"])
            nonce = uuid.uuid4().hex[:10]
            staged_docx = staging_dir / f"wordrev_{safe_job_id}_postprocess_{nonce}.docx"
            staged_docx.unlink(missing_ok=True)
            try:
                # Keep Word reads and writes inside the stable macOS-authorized
                # directory, then copy the refreshed package back atomically.
                shutil.copy2(accepted, staged_docx)
                _disable_automatic_field_refresh(staged_docx)
                refresh = _word_refresh_selected(staged_docx)
                shutil.copy2(staged_docx, accepted)
            finally:
                staged_docx.unlink(missing_ok=True)
        else:
            refresh = _word_refresh_selected(accepted)
        refresh["execution_skipped"] = False
    Path(paths["refresh_log"]).write_text(json.dumps(refresh, ensure_ascii=False, indent=2), encoding="utf-8")


def phase_outputs(phase: str, job: dict, paths: dict[str, str]) -> list[str]:
    if phase == "word-postprocess":
        outputs = [paths["accepted"], paths["refresh_log"]]
        if job.get("output", {}).get("tracked_formatting", False) and Path(paths["tracked"]).is_file():
            outputs.append(paths["tracked"])
        return outputs
    return []


def _extract_office_json(stdout: str) -> dict:
    start = stdout.find("{")
    return json.loads(stdout[start:]) if start >= 0 else {"success": False, "raw": stdout}


def _tracked_revision_audit(path: Path, expected_author: str) -> dict[str, Any]:
    tracked_names = {
        "ins", "del", "moveFrom", "moveTo", "pPrChange", "rPrChange",
        "sectPrChange", "tblPrChange", "trPrChange", "tcPrChange",
    }
    by_type: dict[str, int] = {}
    authors: set[str] = set()
    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            if not name.startswith("word/") or not name.endswith(".xml"):
                continue
            root = etree.fromstring(package.read(name))
            for element in root.iter():
                local = etree.QName(element).localname
                author = element.get(qn("w:author"))
                if local in tracked_names and author is not None:
                    by_type[local] = by_type.get(local, 0) + 1
                    authors.add(author)
    total = sum(by_type.values())
    return {
        "total": total,
        "by_type": dict(sorted(by_type.items())),
        "authors": sorted(authors),
        "expected_author": expected_author,
        "all_pass": total > 0 and authors == {expected_author},
    }


def verify(job: dict, paths: dict[str, str], audit: dict, command_count: int) -> dict:
    accepted = Path(paths["accepted"])
    profile_path, profile = load_profile(job)
    before = json.loads(Path(paths["semantic_before"]).read_text(encoding="utf-8"))
    after = _xml_semantic_snapshot(accepted)
    Path(paths["semantic_after"]).write_text(json.dumps(after, ensure_ascii=False, indent=2), encoding="utf-8")
    semantic_checks = {
        key: before.get(key) == after.get(key)
        for key in SEMANTIC_INVARIANT_KEYS
    }
    validation = _extract_office_json(office("validate", str(accepted), "--json").stdout)
    document = _extract_office_json(office("get", str(accepted), "/", "--depth", "1", "--json").stdout)
    office("close", str(accepted), "--json", check=False)
    rows = _office_results(document)
    fmt = rows[0].get("format", {}) if rows else {}
    format_checks = {
        prop: _format_equivalent(prop, fmt.get(prop), target)
        for prop, target in profile["document_props"].items()
    }
    errors = []
    if not validation.get("success"):
        errors.append("officecli_validate_failed")
    if not all(semantic_checks.values()):
        errors.append("semantic_content_changed")
    failed_format = [key for key, passed in format_checks.items() if not passed]
    if failed_format:
        errors.append("profile_document_props_failed")
    refresh = json.loads(Path(paths["refresh_log"]).read_text(encoding="utf-8"))
    if refresh.get("errors", 0):
        errors.append("word_field_refresh_errors")
    warnings = list(json.loads(Path(paths["metadata"]).read_text(encoding="utf-8")).get("warnings", []))
    tracked_qa = {"skipped": True}
    if job.get("output", {}).get("tracked_formatting", False):
        tracked = Path(paths["tracked"])
        if not tracked.is_file():
            errors.append("tracked_formatting_missing")
        else:
            tracked_qa = _tracked_revision_audit(
                tracked, job.get("revision_author", "Codex")
            )
            tracked_qa["validation"] = _extract_office_json(
                office("validate", str(tracked), "--json").stdout
            )
            office("close", str(tracked), "--json", check=False)
            if not tracked_qa["all_pass"]:
                errors.append("tracked_formatting_revision_audit_failed")
            if not tracked_qa["validation"].get("success"):
                errors.append("tracked_formatting_validation_failed")
    if after.get("floating_drawings", 0):
        warnings.append(f"floating_drawings_require_review:{after['floating_drawings']}")
    result = {
        "all_pass": not errors,
        "profile": profile["profile_id"],
        "profile_sha256": sha256(profile_path),
        "scope": job.get("scope"),
        "source_sha256": sha256(Path(job["source"])),
        "accepted_sha256": sha256(accepted),
        "officecli_commands": command_count,
        "semantic_checks": semantic_checks,
        "format_checks": format_checks,
        "failed_format_properties": failed_format,
        "validation": validation,
        "field_refresh": refresh,
        "tracked_formatting": tracked_qa,
        "warnings": sorted(set(warnings)),
        "errors": errors,
        "notebooklm": "disabled",
        "statistics_executed": False,
    }
    Path(paths["adapter_qa_json"]).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    Path(paths["qa_md"]).write_text(
        "# Thesis format QA\n\n"
        f"- Profile: `{profile['profile_id']}`\n"
        f"- Scope: `{job.get('scope')}`\n"
        f"- Passed: `{result['all_pass']}`\n"
        f"- Semantic checks: `{sum(semantic_checks.values())}/{len(semantic_checks)}`\n"
        f"- Document format checks: `{sum(format_checks.values())}/{len(format_checks)}`\n"
        f"- Warnings: `{'; '.join(result['warnings']) or 'none'}`\n"
        f"- Errors: `{'; '.join(errors) or 'none'}`\n"
        "- `view issues` findings remain a review queue and were not auto-deleted.\n",
        encoding="utf-8",
    )
    if errors:
        raise RuntimeError(f"Thesis format verification failed: {errors}")
    return result


def verify_pdf(job: dict, paths: dict[str, str], pdf: Path, generic_pdf: dict) -> dict[str, Any]:
    page_count = int(generic_pdf.get("page_count") or 1)
    process = run(["pdfinfo", "-f", "1", "-l", str(page_count), "-box", str(pdf)])
    sizes = []
    for line in process.stdout.splitlines():
        match = re.search(r"Page\s+(\d+)\s+size:\s+([0-9.]+)\s+x\s+([0-9.]+)\s+pts", line)
        if match:
            page, width, height = int(match.group(1)), float(match.group(2)), float(match.group(3))
            a4 = (
                (abs(width - 595.28) < 4 and abs(height - 841.89) < 4)
                or (abs(width - 841.89) < 4 and abs(height - 595.28) < 4)
            )
            sizes.append({"page": page, "width_pt": width, "height_pt": height, "a4": a4, "orientation": "landscape" if width > height else "portrait"})
    non_a4 = [row["page"] for row in sizes if not row["a4"]]
    unexpected_blank = generic_pdf.get("unexpected_blank_pages", [])
    result = {
        "all_pass": not non_a4 and not unexpected_blank,
        "page_count": generic_pdf.get("page_count"),
        "page_sizes": sizes,
        "non_a4_pages": non_a4,
        "blank_pages": generic_pdf.get("blank_pages", []),
        "allowed_blank_pages": generic_pdf.get("allowed_blank_pages", []),
        "unexpected_blank_pages": unexpected_blank,
        "manual_review_required": ["first page", "first odd/even body pages", "TOC", "captions", "landscape pages", "footnotes", "references", "last page"],
    }
    return result


def _word_pdf_script() -> str:
    return r'''
on run argv
  set inputName to item 1 of argv
  set outputPath to item 2 of argv
  with timeout of 900 seconds
  tell application "Microsoft Word"
    set previousAlerts to display alerts
    set docRef to missing value
    try
      set display alerts to alerts none
      repeat 240 times
        try
          set docRef to document inputName
          exit repeat
        on error
          delay 0.25
        end try
      end repeat
      if docRef is missing value then error "Word PDF input did not open"
      set show revisions of docRef to false
      set print revisions of docRef to false
      save as docRef file name POSIX file outputPath file format format PDF add to recent files false
      close docRef saving no
      set docRef to missing value
      set display alerts to previousAlerts
    on error errMsg number errNum
      try
        if docRef is not missing value then close docRef saving no
      end try
      set display alerts to previousAlerts
      error errMsg number errNum
    end try
  end tell
  end timeout
end run
'''


def _export_pdf_unlocked(source: Path, output: Path, *, print_markup: bool = False) -> None:
    del print_markup
    run(["open", "-a", "Microsoft Word", str(source)], timeout=60)
    time.sleep(3)
    script = _word_pdf_script()
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile("w", suffix=".applescript", encoding="utf-8", delete=False) as handle:
        handle.write(script)
        script_path = Path(handle.name)
    try:
        run(["osascript", str(script_path), source.name, str(output)], timeout=900)
    finally:
        script_path.unlink(missing_ok=True)
    if not output.is_file():
        raise RuntimeError(f"Microsoft Word did not export PDF: {output}")


def export_pdf(source: Path, output: Path, *, print_markup: bool = False) -> None:
    with _word_process_lock(_word_lock_path(), WORD_LOCK_TIMEOUT_SECONDS):
        _export_pdf_unlocked(source, output, print_markup=print_markup)
