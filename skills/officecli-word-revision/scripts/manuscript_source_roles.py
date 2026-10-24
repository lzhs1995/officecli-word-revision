"""Source-bound exceptions to the uniform manuscript typography contract.

These paragraphs are checked against the immutable source, never exempted
from checking. The sidecar contains selectors, not replacement formatting.
Only the accepted view is indexed; ordinary blank paragraphs remain anchors.
"""
from __future__ import annotations

import copy
import hashlib
import io
import json
import re
from collections import Counter
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROLES = {"abstract", "keywords", "bibliography_entry"}
SHA = re.compile(r"[0-9a-f]{64}")
LABELS = {
    "abstract": re.compile(r"^(?:摘要|abstract)\s*[:：]", re.I),
    "keywords": re.compile(r"^(?:关键词|关键字|key\s*words)\s*[:：]", re.I),
}
BIB_HEADING = re.compile(r"^(?:参考文献|references|bibliography)\s*[:：]?\s*$", re.I)
P_BOOLEAN = {
    "keepNext", "keepLines", "widowControl", "pageBreakBefore", "suppressLineNumbers",
    "bidi", "contextualSpacing", "suppressAutoHyphens", "snapToGrid",
}
REVISION = {qn("w:" + tag) for tag in
            ("rPrChange", "pPrChange", "ins", "del", "moveFrom", "moveTo")}


class SourceRoleError(ValueError):
    pass


def _keys(value, required, optional=()):
    if not isinstance(value, dict) or not set(required) <= value.keys() or (
        value.keys() - set(required) - set(optional)
    ):
        raise SourceRoleError("SOURCE_ROLE_CONTRACT_KEYS_INVALID")


def _index(value):
    if type(value) is not int or value < 0:
        raise SourceRoleError("SOURCE_ROLE_INDEX_INVALID")
    return value


def _digest(text):
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _read_pin(spec):
    _keys(spec, ("path", "bytes", "sha256"))
    if not isinstance(spec["path"], str) or not Path(spec["path"]).is_absolute():
        raise SourceRoleError("SOURCE_ROLE_PATH_MUST_BE_ABSOLUTE")
    if type(spec["bytes"]) is not int or spec["bytes"] <= 0 or (
        not isinstance(spec["sha256"], str) or not SHA.fullmatch(spec["sha256"])
    ):
        raise SourceRoleError("SOURCE_ROLE_PIN_INVALID")
    try:
        raw = Path(spec["path"]).read_bytes()
    except OSError as exc:
        raise SourceRoleError("SOURCE_ROLE_PIN_UNREADABLE") from exc
    if len(raw) != spec["bytes"] or hashlib.sha256(raw).hexdigest() != spec["sha256"]:
        raise SourceRoleError("SOURCE_ROLE_PIN_MISMATCH")
    return raw


def _json(raw):
    def pairs(items):
        result = {}
        for key, value in items:
            if key in result:
                raise SourceRoleError("SOURCE_ROLE_DUPLICATE_JSON_KEY")
            result[key] = value
        return result

    def constant(value):
        raise SourceRoleError("SOURCE_ROLE_INVALID_JSON_NUMBER")

    try:
        return json.loads(raw.decode("utf-8"), object_pairs_hook=pairs,
                          parse_constant=constant)
    except (UnicodeError, json.JSONDecodeError) as exc:
        raise SourceRoleError("SOURCE_ROLE_JSON_INVALID") from exc


def _visible(element):
    from manuscript_format import visible
    return visible(element)


def _tokens(run):
    """Keep tabs, break kinds and literal whitespace distinct."""
    result = []
    for child in run:
        if not _visible(child):
            continue
        tag = child.tag
        if tag == qn("w:t"):
            result.extend(("text", c) for c in child.text or "")
        elif tag in {qn("w:tab"), qn("w:cr"), qn("w:softHyphen"),
                     qn("w:noBreakHyphen")}:
            result.append((tag.split("}")[-1], ""))
        elif tag == qn("w:br"):
            result.append(("br", child.get(qn("w:type"), "textWrapping"),
                           child.get(qn("w:clear"), "none")))
        elif tag == qn("w:sym"):
            result.append(("sym", child.get(qn("w:font"), ""),
                           child.get(qn("w:char"), "")))
    return result


def accepted_paragraphs(document):
    result = []
    for raw_index, paragraph in enumerate(document.element.body.findall(qn("w:p"))):
        if not _visible(paragraph):
            continue
        deleted_mark = paragraph.find("w:pPr/w:rPr/w:del", paragraph.nsmap)
        moved_mark = paragraph.find("w:pPr/w:rPr/w:moveFrom", paragraph.nsmap)
        if deleted_mark is not None or moved_mark is not None:
            if any(_tokens(r) for r in paragraph.xpath(".//w:r") if _visible(r)):
                # A deleted mark with surviving content merges paragraphs in
                # Word. Do not guess that more complicated coordinate mapping.
                raise SourceRoleError("SOURCE_ROLE_PARAGRAPH_MERGE_UNSUPPORTED")
            continue
        result.append((raw_index, paragraph))
    return result


def _ppr_layers(document, paragraph):
    styles = {s.style_id: s.element for s in document.styles}
    yield document.styles.element.find(
        "w:docDefaults/w:pPrDefault/w:pPr", document.styles.element.nsmap)

    def chain(sid, seen):
        if not sid or sid in seen or sid not in styles:
            return
        seen.add(sid)
        style = styles[sid]
        base = style.find(qn("w:basedOn"))
        if base is not None:
            yield from chain(base.get(qn("w:val")), seen)
        yield style.find(qn("w:pPr"))

    from manuscript_format import style_id
    default = next((s.style_id for s in document.styles
                    if s.element.get(qn("w:default")) == "1"
                    and s.element.get(qn("w:type")) == "paragraph"), None)
    yield from chain(style_id(paragraph) or default, set())
    yield paragraph.find(qn("w:pPr"))


def _paragraph_state(document, paragraph):
    from manuscript_format import effective_paragraph
    value = copy.deepcopy(effective_paragraph(document, paragraph))
    # A removed redundant left alignment, or true/1 boolean spelling, is the
    # same effective format. Missing line/lineRule is deliberately NOT invented.
    value.setdefault("jc", {"val": "left"})
    for key in P_BOOLEAN & value.keys():
        value[key]["val"] = value[key].get("val", "1").lower() not in {"0", "false", "off"}
    stops = {}
    for pr in _ppr_layers(document, paragraph):
        if pr is None:
            continue
        tabs = pr.find(qn("w:tabs"))
        if tabs is not None:
            for tab in tabs:
                attrs = tuple(sorted((k.split("}")[-1], v) for k, v in tab.attrib.items()))
                position = tab.get(qn("w:pos"))
                if tab.get(qn("w:val")) == "clear":
                    stops.pop(position, None)
                else:
                    stops[position] = attrs
    if stops or "tabs" in value:
        value["tabs"] = sorted(stops.values())
    return value


def _mark_state(document, paragraph):
    from manuscript_format import effective_run
    run = OxmlElement("w:r")
    mark = paragraph.find("w:pPr/w:rPr", paragraph.nsmap)
    if mark is not None:
        run.append(copy.deepcopy(mark))
    return effective_run(document, paragraph, run)


def _stream(document, paragraph):
    from manuscript_format import effective_run
    result = []
    for run in paragraph.xpath(".//w:r"):
        if _visible(run):
            tokens = _tokens(run)
            if tokens:
                result.append({"run": run, "tokens": tokens,
                               "format": effective_run(document, paragraph, run)})
    return result


def _flatten(stream):
    return [token for segment in stream for token in segment["tokens"]]


def _overlaps(expected, actual):
    """Compare positions, not XML run ordinal or just a font multiset."""
    ei = ai = eoff = aoff = position = 0
    while ei < len(expected) and ai < len(actual):
        n = min(len(expected[ei]["tokens"]) - eoff, len(actual[ai]["tokens"]) - aoff)
        yield position, position + n, expected[ei], actual[ai]
        position += n
        eoff += n
        aoff += n
        if eoff == len(expected[ei]["tokens"]):
            ei += 1
            eoff = 0
        if aoff == len(actual[ai]["tokens"]):
            ai += 1
            aoff = 0


def _check_source(contract, document):
    _keys(contract, ("schema_version", "source", "paragraph_count", "entries"),
          ("bibliography",))
    if contract["schema_version"] != "1.0":
        raise SourceRoleError("SOURCE_ROLE_SCHEMA_VERSION_INVALID")
    paragraphs = accepted_paragraphs(document)
    if _index(contract["paragraph_count"]) != len(paragraphs):
        raise SourceRoleError("SOURCE_ROLE_PARAGRAPH_COUNT_MISMATCH")
    entries = contract["entries"]
    if not isinstance(entries, list) or not entries:
        raise SourceRoleError("SOURCE_ROLE_SELECTION_EMPTY")
    from manuscript_format import paragraph_role, style_id, text
    selections = {}
    for entry in entries:
        _keys(entry, ("index", "role", "text_sha256", "style_id"))
        index = _index(entry["index"])
        if index in selections or index >= len(paragraphs):
            raise SourceRoleError("SOURCE_ROLE_DUPLICATE_OR_OUT_OF_RANGE")
        role = entry["role"]
        if not isinstance(role, str) or role not in ROLES:
            raise SourceRoleError("SOURCE_ROLE_UNKNOWN")
        p = paragraphs[index][1]
        content = text(p)
        if not content.strip() or _digest(content) != entry["text_sha256"]:
            raise SourceRoleError("SOURCE_ROLE_SOURCE_TEXT_MISMATCH")
        if style_id(p) != entry["style_id"]:
            raise SourceRoleError("SOURCE_ROLE_SOURCE_STYLE_MISMATCH")
        if paragraph_role(document, p) in {"caption", "image", "equation", "heading", "note"}:
            raise SourceRoleError("SOURCE_ROLE_CANNOT_RECLASSIFY_SHARED_ROLE")
        if role in LABELS and not LABELS[role].match(content.strip()):
            raise SourceRoleError("SOURCE_ROLE_LABEL_MISMATCH")
        selections[index] = entry
    for index, (_, paragraph) in enumerate(paragraphs):
        content = text(paragraph).strip()
        for role, label in LABELS.items():
            if label.match(content) and selections.get(index, {}).get("role") != role:
                raise SourceRoleError("SOURCE_ROLE_LABEL_COVERAGE_MISMATCH")
    bib = {i for i, e in selections.items() if e["role"] == "bibliography_entry"}
    headings = [i for i, (_, p) in enumerate(paragraphs) if BIB_HEADING.fullmatch(text(p).strip())]
    if bib:
        bounds = contract.get("bibliography")
        _keys(bounds, ("heading_index", "end_index"))
        heading, end = _index(bounds["heading_index"]), _index(bounds["end_index"])
        if headings != [heading] or not heading < end <= len(paragraphs) or not BIB_HEADING.fullmatch(
            text(paragraphs[heading][1]).strip()
        ):
            raise SourceRoleError("SOURCE_ROLE_BIBLIOGRAPHY_BOUNDARY_INVALID")
        expected = {i for i in range(heading + 1, end) if text(paragraphs[i][1]).strip()}
        if not expected or bib != expected:
            raise SourceRoleError("SOURCE_ROLE_BIBLIOGRAPHY_COVERAGE_MISMATCH")
        if any(i in selections and i not in bib for i in range(heading, end)):
            raise SourceRoleError("SOURCE_ROLE_BIBLIOGRAPHY_ROLE_CONFLICT")
    elif "bibliography" in contract or headings:
        raise SourceRoleError("SOURCE_ROLE_BIBLIOGRAPHY_SELECTION_MISSING")
    return paragraphs, selections


def load_bound_source(pointer, *, source_path=None):
    raw = _read_pin(pointer)
    contract = _json(raw)
    _keys(contract, ("schema_version", "source", "paragraph_count", "entries"), ("bibliography",))
    source_raw = _read_pin(contract["source"])
    if source_path is not None and Path(contract["source"]["path"]).resolve() != Path(source_path).resolve():
        raise SourceRoleError("SOURCE_ROLE_CONTENT_SOURCE_PATH_MISMATCH")
    try:
        document = Document(io.BytesIO(source_raw))
    except Exception as exc:
        raise SourceRoleError("SOURCE_ROLE_SOURCE_DOCX_INVALID") from exc
    paragraphs, selections = _check_source(contract, document)
    return contract, document, paragraphs, selections


class Binding:
    def __init__(self, pointer, document):
        self.contract, self.source, source_paragraphs, selections = load_bound_source(pointer)
        self.pointer = pointer
        actual = accepted_paragraphs(document)
        if len(actual) != len(source_paragraphs):
            raise SourceRoleError("SOURCE_ROLE_ACCEPTED_VIEW_COUNT_MISMATCH")
        self.targets = {}
        for index, entry in selections.items():
            raw_index, paragraph = actual[index]
            source_p = source_paragraphs[index][1]
            self.targets[raw_index] = {
                "entry": entry, "paragraph": paragraph, "source": source_p,
                "stream": _stream(self.source, source_p),
                "paragraph_format": _paragraph_state(self.source, source_p),
                "mark_format": _mark_state(self.source, source_p),
            }

    def audit(self, document):
        from manuscript_format import text
        errors, checked = [], []
        coverage = True
        for raw_index, target in self.targets.items():
            p = target["paragraph"]
            role = target["entry"]["role"]
            loc = f"p:{raw_index}:{role}"
            actual = _stream(document, p)
            if text(p) != text(target["source"]) or _flatten(actual) != _flatten(target["stream"]):
                errors.append({"location": loc, "rule": "source_role_text_or_controls"})
                coverage = False
            else:
                for start, end, want, got in _overlaps(target["stream"], actual):
                    if want["format"] != got["format"]:
                        errors.append({"location": loc, "rule": "source_role_run_format",
                                       "start": start, "end": end,
                                       "expected": want["format"], "actual": got["format"]})
            for rule, want, got in [
                ("source_role_paragraph_format", target["paragraph_format"], _paragraph_state(document, p)),
                ("source_role_paragraph_mark", target["mark_format"], _mark_state(document, p)),
            ]:
                if want != got:
                    errors.append({"location": loc, "rule": rule, "expected": want, "actual": got})
            checked.append({"location": loc, "role": role, "index": target["entry"]["index"],
                            "text_runs": sum(bool(_tokens(r)) for r in p.xpath(".//w:r") if _visible(r))})
        return {"all_pass": not errors and coverage, "coverage_ok": coverage,
                "source": self.contract["source"], "sidecar": self.pointer,
                "roles": dict(Counter(r["role"] for r in checked)),
                "paragraphs_checked": len(checked), "checked": checked, "errors": errors}

    def apply(self, document, recorder):
        from manuscript_format import effective_run, text
        for raw_index, target in self.targets.items():
            p, source_p = target["paragraph"], target["source"]
            actual = _stream(document, p)
            if text(p) != text(source_p) or _flatten(actual) != _flatten(target["stream"]):
                raise SourceRoleError("SOURCE_ROLE_APPLY_TEXT_OR_CONTROLS_MISMATCH")
            if _paragraph_state(document, p) != target["paragraph_format"] or (
                _mark_state(document, p) != target["mark_format"]
            ):
                reference = source_p.find(qn("w:pPr"))

                def restore(pr):
                    for child in list(pr):
                        if child.tag not in REVISION | {qn("w:sectPr"), qn("w:rPr")}:
                            pr.remove(child)
                    if reference is not None:
                        for child in reference:
                            if child.tag not in REVISION | {qn("w:sectPr"), qn("w:rPr")}:
                                pr.append(copy.deepcopy(child))
                    mark = pr.find(qn("w:rPr"))
                    original = reference.find(qn("w:rPr")) if reference is not None else None
                    if mark is None and original is not None:
                        mark = OxmlElement("w:rPr")
                        pr.append(mark)
                    if mark is not None:
                        _restore_rpr(mark, original)

                recorder.edit(p, "w:pPr", restore, f"p:{raw_index}:source_role")
            # Equivalent segmentation is accepted. A merged run that crosses
            # distinct source formats needs safe text/control-only splitting.
            for segment in list(actual):
                overlaps = []
                for start, end, want, got in _overlaps(target["stream"], actual):
                    if got is segment:
                        overlaps.append((end - start, want))
                groups = []
                for n, want in overlaps:
                    if groups and groups[-1][1]["format"] == want["format"]:
                        groups[-1] = (groups[-1][0] + n, groups[-1][1])
                    else:
                        groups.append((n, want))
                run = segment["run"]
                if len(groups) > 1:
                    runs = _split_run(run, segment["tokens"], [n for n, _ in groups])
                else:
                    runs = [run]
                for piece, (_, want) in zip(runs, groups):
                    if effective_run(document, p, piece) != want["format"]:
                        reference = want["run"].find(qn("w:rPr"))
                        recorder.edit(piece, "w:rPr", lambda pr, ref=reference: _restore_rpr(pr, ref),
                                      f"p:{raw_index}:source_run")
        result = self.audit(document)
        if result["errors"]:
            raise SourceRoleError("SOURCE_ROLE_APPLY_UNRESOLVED_FORMAT")


def _restore_rpr(pr, reference):
    for child in list(pr):
        if child.tag not in REVISION:
            pr.remove(child)
    if reference is not None:
        for child in reference:
            if child.tag not in REVISION:
                pr.append(copy.deepcopy(child))


def _split_run(run, tokens, sizes):
    allowed = {qn("w:" + tag) for tag in
               ("rPr", "t", "tab", "br", "cr", "sym", "softHyphen", "noBreakHyphen")}
    if any(child.tag not in allowed for child in run) or run.xpath("./w:rPr/w:rPrChange"):
        raise SourceRoleError("SOURCE_ROLE_RUN_SPLIT_UNSUPPORTED")
    result, offset = [], 0
    for size in sizes:
        piece = copy.deepcopy(run)
        for child in list(piece):
            if child.tag != qn("w:rPr"):
                piece.remove(child)
        for token in tokens[offset:offset + size]:
            kind = token[0]
            if kind == "text" and len(piece) and piece[-1].tag == qn("w:t"):
                piece[-1].text = (piece[-1].text or "") + token[1]
            else:
                child = OxmlElement("w:t" if kind == "text" else "w:" + kind)
                if kind == "text":
                    child.text = token[1]
                    child.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                elif kind == "br":
                    child.set(qn("w:type"), token[1])
                    child.set(qn("w:clear"), token[2])
                elif kind == "sym":
                    child.set(qn("w:font"), token[1])
                    child.set(qn("w:char"), token[2])
                piece.append(child)
        offset += size
        run.addprevious(piece)
        result.append(piece)
    run.getparent().remove(run)
    return result
