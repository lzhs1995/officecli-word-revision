"""Compare saved Word field caches without mistaking regenerated TOC names for drift.

This is a read-only comparison of supplied saved files and observed page counts.
It does not run Word or establish that the caller actually refreshed any field.
"""
from __future__ import annotations

import hashlib
import json
import re

from docx.oxml.ns import qn

from manuscript_format import visible
from thesis_integrity import Audit, digest

KINDS = {"TOC", "SEQ", "REF", "PAGEREF", "PAGE", "NUMPAGES"}
AUTO_TOC = re.compile(r"_Toc[0-9]+\Z")


def snapshot(path, contract, pages):
    if not isinstance(pages, int) or isinstance(pages, bool) or pages < 1:
        raise ValueError("An observed positive page count is required")
    a = Audit(path, contract)
    a.scan()
    errors = list(a.errors)
    positions = {}
    for part, root in a.parts.items():
        cursor = 0
        positions[part] = {}

        def locate(node):
            nonlocal cursor
            positions[part][node] = cursor
            if not visible(node):
                return
            if node.tag in {qn("w:t"), qn("m:t")}:
                cursor += len(node.text or "")
            elif node.tag in {qn("w:tab"), qn("w:br"), qn("w:drawing"), qn("w:pict"), qn("w:object")}:
                cursor += 1
            for child in node:
                locate(child)
            if node.tag == qn("w:p"):
                cursor += 1

        # Display-character offsets plus paragraph/object boundaries survive
        # Word's run splitting and rPr materialization. XML-node offsets do not.
        locate(root)
    resolved = {}

    def target(name):
        if not AUTO_TOC.fullmatch(name):
            return name
        row = a.bookmarks.get(name)
        if row is None or not row["text"].strip():
            errors.append({"code": "AUTO_TOC_TARGET_UNRESOLVED", "name": name})
            return name
        value = {"part": row["part"], "start": positions[row["part"]][row["node"]],
                 "end": positions[row["part"]][row["end"]], "text": row["text"]}
        # Position and text both participate: identical titles at different
        # places must not make a misdirected cross-reference look stable.
        fingerprint = hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True).encode()).hexdigest()
        resolved[name] = {**value, "fingerprint": fingerprint}
        return "AUTO_TOC_TARGET:" + fingerprint

    raw_fields, semantic_fields = [], []
    for f in a.fields:
        kind = a.field_kind(f)
        if kind not in KINDS:
            continue
        code = f["instruction"]
        raw = {"part": f["part"], "kind": kind, "code": code, "result": f["result"]}
        raw_fields.append(raw)
        match = re.match(r'\s*(?:REF|PAGEREF)\s+("[^"]+"|\S+)', code, re.I)
        if match:
            name = match[1].strip('"')
            if AUTO_TOC.fullmatch(name):
                code = code[:match.start(1)] + target(name) + code[match.end(1):]
        semantic_fields.append({**raw, "code": code})
    raw_links, semantic_links = [], []
    for part, root in a.parts.items():
        for node in root.iter(qn("w:hyperlink")):
            if visible(node) and node.get(qn("w:anchor")):
                name = node.get(qn("w:anchor"))
                value = {"part": part, "position": positions[part][node], "anchor": name}
                raw_links.append(value)
                semantic_links.append({**value, "anchor": target(name)})
    return {"schema_version": "1.0", "docx_sha256": digest(path),
            "raw": {"pages": pages, "fields": raw_fields, "internal_links": raw_links},
            "semantic": {"pages": pages, "fields": semantic_fields, "internal_links": semantic_links},
            "resolved_auto_toc_targets": resolved, "errors": errors,
            "scope": "observed page count, selected field code/cache and internal hyperlink targets; no native execution claim"}


def compare(first, second):
    errors = first["errors"] + second["errors"]
    raw_equal = first["raw"] == second["raw"]
    equal = not errors and first["semantic"] == second["semantic"]
    return {"status": "STABLE" if equal else "REVIEW_REQUIRED",
            "semantic_equal": equal, "raw_equal": raw_equal, "errors": errors,
            "auto_toc_name_only_difference": equal and not raw_equal,
            "first_docx_sha256": first["docx_sha256"], "second_docx_sha256": second["docx_sha256"],
            "first_auto_toc_targets": first["resolved_auto_toc_targets"],
            "second_auto_toc_targets": second["resolved_auto_toc_targets"]}
