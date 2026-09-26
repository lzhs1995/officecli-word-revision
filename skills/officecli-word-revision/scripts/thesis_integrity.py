"""Read-only, explicitly scoped whole-thesis OOXML checks.

This is structural evidence, not Word pagination or semantic acceptance. The
contract names objects and policy; the auditor never guesses and rewrites them.
"""
from __future__ import annotations

import argparse
from collections import Counter
import hashlib
import json
from pathlib import Path
import re
import unicodedata
import zipfile

from docx import Document
from docx.oxml.ns import qn
from lxml import etree
from manuscript_format import effective_paragraph, effective_run, visible

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"}
SCHEMA = Path(__file__).resolve().parent.parent / "references/thesis-integrity.schema.json"
FALSE = {"0", "false", "off"}
HAN = re.compile(r"[\u3400-\u9fff\uf900-\ufaff\U00020000-\U000323af]")


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def normalized(value):
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", value))


def text(node):
    return "".join(n.text or "" for n in node.iter()
                   if n.tag in {qn("w:t"), qn("m:t")} and visible(n))


def on(value):
    return str(value).lower() not in FALSE


def xpath(node, expression):
    return etree.XPath(expression, namespaces=NS)(node)


def paragraph_toggle(props, key):
    return key in props and on(props[key].get("val", "true"))


def validate_contract(spec):
    from jsonschema import Draft202012Validator
    Draft202012Validator(json.loads(SCHEMA.read_text(encoding="utf-8"))).validate(spec)
    for key in ("blocks", "captions"):
        ids = [r["id"] for r in spec.get(key, [])]
        if len(ids) != len(set(ids)):
            raise ValueError(f"Duplicate {key} IDs")
    entries = spec.get("bibliography", {}).get("entries", [])
    if len({r["id"] for r in entries}) != len(entries):
        raise ValueError("Duplicate bibliography IDs")


def load_bound_contract(binding, *, validate=True):
    if not isinstance(binding, dict) or set(binding) != {"path", "sha256"}:
        raise ValueError("QA binding requires exactly path and sha256")
    path = Path(binding["path"])
    if not path.is_absolute() or not re.fullmatch(r"[a-f0-9]{64}", binding["sha256"]):
        raise ValueError("QA binding requires an absolute path and SHA256")
    payload = path.read_bytes()
    if hashlib.sha256(payload).hexdigest() != binding["sha256"]:
        raise ValueError("QA_CONTRACT_CHANGED: " + str(path))
    spec = json.loads(payload)
    if validate:
        validate_contract(spec)
    return spec


def parse_fields(root, part):
    """Respect nested, adjacent, simple and paragraph-spanning fields per story.

    Text boxes and individual notes have separate stacks. Deleted text never
    repairs an unbalanced live field. Original instruction strings are retained.
    """
    fields, errors, stack = [], [], []

    def issue(code):
        errors.append({"code": code, "part": part})

    def start(node, instruction="", simple=False):
        stack.append({"node": node, "part": part, "instruction": instruction,
                      "result": "", "result_nodes": [], "separated": simple, "simple": simple,
                      "locked": on(node.get(qn("w:fldLock"), "false"))})

    def walk(node, is_root=False):
        if not visible(node):
            return
        if not is_root and node.tag in {qn("w:txbxContent"), qn("w:footnote"), qn("w:endnote")}:
            subfields, suberrors = parse_fields(node, part)
            fields.extend(subfields)
            errors.extend(suberrors)
            return
        simple = node.tag == qn("w:fldSimple")
        if simple:
            start(node, node.get(qn("w:instr"), ""), True)
        elif node.tag == qn("w:fldChar"):
            kind = node.get(qn("w:fldCharType"))
            if kind == "begin":
                start(node)
            elif kind == "separate":
                if not stack or stack[-1]["simple"] or stack[-1]["separated"]:
                    issue("FIELD_UNEXPECTED_SEPARATOR")
                else:
                    stack[-1]["separated"] = True
            elif kind == "end":
                if not stack or stack[-1]["simple"]:
                    issue("FIELD_UNEXPECTED_END")
                else:
                    fields.append(stack.pop())
        elif node.tag == qn("w:instrText"):
            if stack and not stack[-1]["separated"]:
                stack[-1]["instruction"] += node.text or ""
            else:
                issue("FIELD_INSTRUCTION_OUTSIDE_CODE")
        elif node.tag in {qn("w:t"), qn("w:tab"), qn("w:br")}:
            value = (node.text or "") if node.tag == qn("w:t") else "\t" if node.tag == qn("w:tab") else "\n"
            for field in stack:
                if field["separated"]:
                    field["result"] += value
                    field["result_nodes"].append(node)
        for child in node:
            walk(child)
        if node.tag == qn("w:p"):
            for field in stack:
                if field["separated"]:
                    field["result"] += "\n"
        if simple:
            if stack and stack[-1]["node"] is node:
                fields.append(stack.pop())
            else:
                issue("SIMPLE_FIELD_UNBALANCED")

    walk(root, True)
    errors.extend({"code": "FIELD_UNCLOSED", "part": part,
                   "instruction": f["instruction"]} for f in stack)
    return fields, errors


class Audit:
    def __init__(self, path, spec):
        validate_contract(spec)
        self.path, self.spec = Path(path), spec
        self.document = Document(path)
        self.parts = {"word/document.xml": self.document.element}
        with zipfile.ZipFile(path) as archive:
            for name in sorted(archive.namelist()):
                if re.fullmatch(r"word/(?:footnotes|endnotes|header[^/]*|footer[^/]*)\.xml", name):
                    self.parts[name] = etree.fromstring(archive.read(name))
        self.errors, self.reviews, self.observations = [], [], {}
        self.body = self.document.element.body
        self.children = list(self.body)
        self.child_index = {node: i for i, node in enumerate(self.children)}
        self.nodes = {name: list(root.iter()) for name, root in self.parts.items()}
        self.positions = {name: {node: i for i, node in enumerate(nodes)} for name, nodes in self.nodes.items()}
        self.bookmarks = {}
        self.fields = []
        self.sections = []
        for i, node in enumerate(self.children):
            sect = node if node.tag == qn("w:sectPr") else node.find("w:pPr/w:sectPr", NS)
            if sect is not None:
                self.sections.append((i, sect))

    def error(self, code, **detail):
        self.errors.append({"code": code, **detail})

    def review(self, code, **detail):
        self.reviews.append({"code": code, **detail})

    def top(self, node):
        while node is not None and node.getparent() is not self.body:
            node = node.getparent()
        return node

    def index(self, node):
        return self.child_index.get(self.top(node), -1)

    def paragraph(self, node):
        return next((n for n in [node, *node.iterancestors()] if n.tag == qn("w:p")), None)

    def section_number(self, node):
        index = self.index(node)
        return next((i + 1 for i, (end, _) in enumerate(self.sections) if 0 <= index <= end), None)

    def scan(self):
        for part, root in self.parts.items():
            fields, errors = parse_fields(root, part)
            self.fields.extend(fields)
            self.errors.extend(errors)
            ends = {}
            for node in root.iter(qn("w:bookmarkEnd")):
                if visible(node):
                    ends.setdefault(node.get(qn("w:id")), []).append(node)
            for node in root.iter(qn("w:bookmarkStart")):
                if not visible(node):
                    continue
                name, ident = node.get(qn("w:name")), node.get(qn("w:id"))
                matches = ends.get(ident, [])
                if not name or name in self.bookmarks:
                    self.error("BOOKMARK_NAME_DUPLICATE_OR_EMPTY", part=part, name=name)
                    continue
                if len(matches) != 1 or self.positions[part][matches[0]] < self.positions[part][node]:
                    self.error("BOOKMARK_RANGE_INVALID", part=part, name=name)
                    continue
                start, end = self.positions[part][node], self.positions[part][matches[0]]
                selected = self.nodes[part][start:end + 1]
                content = "".join(n.text or "" for n in selected if n.tag == qn("w:t") and visible(n))
                self.bookmarks[name] = {"node": node, "end": matches[0], "part": part,
                                        "nodes": selected, "text": content}
        self.observations["stories"] = sorted(self.parts)
        self.observations["field_counts"] = dict(Counter(self.field_kind(f) for f in self.fields))

    @staticmethod
    def field_kind(field):
        code = field["instruction"].strip()
        if re.match(r"ADDIN\s+ZOTERO_ITEM\b", code, re.I):
            return "ZOTERO_ITEM"
        if re.match(r"ADDIN\s+ZOTERO_BIBL\b", code, re.I):
            return "ZOTERO_BIBL"
        return code.split()[0].upper() if code else "EMPTY"

    def anchor(self, value):
        if "bookmark" in value:
            row = self.bookmarks.get(value["bookmark"])
            if row is None or row["part"] != "word/document.xml":
                self.error("ANCHOR_NOT_UNIQUE_IN_BODY", anchor=value)
                return None
            return row["node"]
        matches = [p for p in self.body.iter(qn("w:p")) if visible(p) and normalized(text(p)) == normalized(value["text"])]
        if len(matches) != 1:
            self.error("ANCHOR_NOT_UNIQUE_IN_BODY", anchor=value, matches=len(matches))
            return None
        return matches[0]

    def begins_section(self, node):
        number = self.section_number(node)
        if number is None:
            return False
        start = self.sections[number - 2][0] + 1 if number > 1 else 0
        return all(not text(p).strip() and p.tag != qn("w:tbl") and not xpath(p, ".//w:drawing|.//w:tbl")
                   for p in self.children[start:self.index(node)])

    def begins_page(self, node):
        p = self.paragraph(node)
        if p is not None and paragraph_toggle(effective_paragraph(self.document, p), "pageBreakBefore"):
            return True
        number = self.section_number(node)
        if self.begins_section(node) and number:
            kind = self.sections[number - 1][1].find("w:type", NS)
            if kind is None or kind.get(qn("w:val")) in {"nextPage", "oddPage", "evenPage"}:
                return True
        # An explicit page break before the anchor (not a cached render break).
        for candidate in reversed(self.children[:self.index(node) + 1]):
            tokens = [n for n in candidate.iter() if n.tag in {qn("w:t"), qn("w:br")} and visible(n)]
            if candidate is self.top(node):
                for token in tokens:
                    if token.tag == qn("w:t") and (token.text or "").strip():
                        break
                    if token.tag == qn("w:br") and token.get(qn("w:type")) == "page":
                        return True
            else:
                for token in reversed(tokens):
                    if token.tag == qn("w:br") and token.get(qn("w:type")) == "page":
                        return True
                    if token.tag == qn("w:t") and (token.text or "").strip():
                        return False
        return False

    def numbering(self, section, rule, location):
        pg = section.find("w:pgNumType", NS)
        actual = pg.get(qn("w:fmt"), "decimal") if pg is not None else "decimal"
        wanted = rule.get("format")
        if wanted and actual not in ({"upperRoman", "lowerRoman"} if wanted == "roman" else {wanted}):
            self.error("PAGE_NUMBER_FORMAT", location=location, actual=actual, expected=wanted)
        start = pg.get(qn("w:start")) if pg is not None else None
        if "start" in rule and ((rule["start"] == "continue" and start is not None)
                                  or (rule["start"] != "continue" and start != str(rule["start"]))):
            self.error("PAGE_NUMBER_RESTART", location=location, actual=start, expected=rule["start"])

    def blocks(self):
        rows = []
        for rule in self.spec["blocks"]:
            node = self.anchor(rule["anchor"])
            if node is None:
                continue
            index, number = self.index(node), self.section_number(node)
            rows.append({"id": rule["id"], "body_index": index, "section": number})
            boundary = rule.get("start", "any")
            if boundary == "new_section" and not self.begins_section(node):
                self.error("BLOCK_SECTION_BOUNDARY", block=rule["id"])
            if boundary == "new_page" and not self.begins_page(node):
                self.error("BLOCK_PAGE_BOUNDARY", block=rule["id"])
            if "section" in rule and number != rule["section"]:
                self.error("BLOCK_SECTION", block=rule["id"], actual=number, expected=rule["section"])
            if "page_number" in rule and number:
                self.numbering(self.sections[number - 1][1], rule["page_number"], rule["id"])
        indexes = [r["body_index"] for r in rows]
        if indexes != sorted(set(indexes)):
            self.error("BLOCK_ORDER_OR_OVERLAP", blocks=rows)
        self.observations["blocks"] = rows
        previous = {}
        section_rules = {r["section"]: r for r in self.spec.get("sections", [])}
        if len(section_rules) != len(self.spec.get("sections", [])):
            self.error("DUPLICATE_SECTION_RULE")
        for number in section_rules:
            if number > len(self.sections):
                self.error("SECTION_MISSING", section=number)
        relationships = self.document.part.rels
        for number, (_, sect) in enumerate(self.sections, 1):
            rule = section_rules.get(number, {})
            if "page_number" in rule:
                self.numbering(sect, rule["page_number"], f"section:{number}")
            kind = sect.find("w:type", NS)
            actual_type = kind.get(qn("w:val")) if kind is not None else "nextPage"
            if "start_type" in rule and rule["start_type"] != actual_type:
                self.error("SECTION_START_TYPE", section=number, actual=actual_type)
            size = sect.find("w:pgSz", NS)
            orientation = size.get(qn("w:orient"), "portrait") if size is not None else "portrait"
            if "orientation" in rule and orientation != rule["orientation"]:
                self.error("SECTION_ORIENTATION", section=number, actual=orientation)
            title = sect.find("w:titlePg", NS)
            title_page = title is not None and on(title.get(qn("w:val"), "true"))
            if "title_page" in rule and title_page != rule["title_page"]:
                self.error("SECTION_FIRST_PAGE", section=number, actual=title_page)
            for story in ("header", "footer"):
                for subtype in ("default", "even", "first"):
                    key = story + "_" + subtype
                    refs = [r for r in sect.findall("w:" + story + "Reference", NS) if r.get(qn("w:type")) == subtype]
                    target = previous.get(key)
                    if len(refs) > 1:
                        self.error("STORY_REFERENCE_DUPLICATE", section=number, story=key)
                    if refs:
                        rel = relationships.get(refs[0].get(qn("r:id")))
                        if rel is None or rel.is_external:
                            self.error("STORY_RELATIONSHIP_BROKEN", section=number, story=key)
                        else:
                            target = rel.target_part
                    expected = rule.get("stories", {}).get(key)
                    linked = number > 1 and (not refs or target is previous.get(key))
                    if expected == "independent" and linked or expected == "linked" and not linked:
                        self.error("STORY_LINK", section=number, story=key, expected=expected)
                    if expected == "empty" and target is not None:
                        root = etree.fromstring(target.blob)
                        if text(root).strip() or xpath(root, ".//w:fldChar|.//w:fldSimple|.//w:drawing"):
                            self.error("STORY_NOT_EMPTY", section=number, story=key)
                    previous[key] = target
        self.observations["section_count"] = len(self.sections)

    def fields_check(self):
        policy = self.spec.get("fields", {})
        by_kind = {}
        for field in self.fields:
            kind = self.field_kind(field)
            by_kind.setdefault(kind, []).append(field)
            code = field["instruction"]
            if policy.get("unlocked") and field["locked"]:
                self.error("FIELD_LOCKED", part=field["part"], instruction=code)
            if kind in {"REF", "PAGEREF", "NOTEREF"}:
                match = re.match(r'\s*\w+\s+(?:"([^"]+)"|([^\s\\]+))', code)
                target = next((v for v in match.groups() if v is not None), None) if match else None
                bookmark = self.bookmarks.get(target)
                has_payload = bookmark is not None and (normalized(bookmark["text"]) or any(n.tag in {qn("w:drawing"), qn("w:footnoteReference"), qn("w:endnoteReference")} for n in bookmark["nodes"]))
                # PAGEREF legitimately targets a zero-width positional bookmark.
                if bookmark is None or (kind == "REF" and not has_payload):
                    self.error("REFERENCE_TARGET_MISSING_OR_EMPTY", part=field["part"], target=target)
            if re.search(r"Error!\s*(Reference source not found|Bookmark not defined)|错误[！!].*(引用源|书签)", field["result"], re.I):
                self.error("FIELD_VISIBLE_ERROR", part=field["part"], instruction=code)
            if kind == "ZOTERO_ITEM":
                try:
                    data, _ = json.JSONDecoder().raw_decode(code[code.index("{"):])
                    if not isinstance(data, dict) or not isinstance(data.get("properties", {}), dict):
                        raise ValueError("invalid citation object/properties")
                    items = data["citationItems"]
                    if not isinstance(items, list) or not items:
                        raise ValueError("no citation items")
                    uris = []
                    for item in items:
                        if not isinstance(item, dict):
                            raise ValueError("citation item must be an object")
                        values = item.get("uris")
                        if not isinstance(values, list) or not values or not all(isinstance(u, str) and u for u in values):
                            raise ValueError("missing item URI")
                        uris.extend(values)
                        if item.get("dontUpdate"):
                            self.error("ZOTERO_DONT_UPDATE", part=field["part"])
                    if len(uris) != len(set(uris)):
                        self.error("ZOTERO_DUPLICATE_URI_WITHIN_FIELD", part=field["part"])
                    if data.get("properties", {}).get("dontUpdate"):
                        self.error("ZOTERO_DONT_UPDATE", part=field["part"])
                except (ValueError, KeyError, TypeError) as exc:
                    self.error("ZOTERO_ITEM_INVALID", part=field["part"], detail=str(exc))
            if kind == "ZOTERO_BIBL":
                try:
                    data, _ = json.JSONDecoder().raw_decode(code[code.index("{"):])
                    if not isinstance(data, dict) or "CSL_BIBLIOGRAPHY" not in code:
                        raise ValueError("invalid bibliography code")
                    if any(key in data and not isinstance(data[key], list) for key in ("uncited", "omitted", "custom")):
                        raise ValueError("invalid bibliography metadata lists")
                    if not normalized(field["result"]):
                        raise ValueError("empty bibliography result")
                except (ValueError, TypeError) as exc:
                    self.error("ZOTERO_BIBLIOGRAPHY_INVALID", part=field["part"], detail=str(exc))
        toc = policy.get("toc")
        if toc:
            counts = {"contents": 0, "figures": 0, "tables": 0, "other": 0}
            for field in by_kind.get("TOC", []):
                label = re.search(r'\\c\s+(?:"([^"]+)"|(\S+))', field["instruction"], re.I)
                label = (label[1] or label[2]) if label else None
                group = "contents" if label is None else "figures" if label == toc["figure_label"] else "tables" if label == toc["table_label"] else "other"
                selector = re.search(r'\\t\s+(?:"([^"]+)"|(\S+))', field["instruction"], re.I)
                if selector:
                    values = [v.strip() for v in (selector[1] or selector[2]).split(",")]
                    if len(values) % 2 or any(not re.fullmatch(r"[1-9]", level) for level in values[1::2]):
                        self.error("TOC_STYLE_SELECTOR_INVALID", instruction=field["instruction"])
                    styles = set(values[::2])
                    matched = [kind for kind, key in (("figures", "figure_styles"), ("tables", "table_styles"))
                               if styles & set(toc.get(key, []))]
                    if len(matched) > 1 or (label is not None and matched and matched != [group]):
                        self.error("TOC_AMBIGUOUS_SELECTOR", instruction=field["instruction"])
                        group = "other"
                    elif matched:
                        group = matched[0]
                counts[group] += 1
                if not normalized(field["result"]):
                    self.error("TOC_EMPTY_CACHE", group=group)
            for group in ("contents", "figures", "tables"):
                if counts[group] != toc[group]:
                    self.error("TOC_COUNT", group=group, actual=counts[group], expected=toc[group])
            self.observations["toc"] = counts
        if len(by_kind.get("ZOTERO_ITEM", [])) < policy.get("minimum_zotero_items", 0):
            self.error("ZOTERO_ITEM_COUNT")
        if "zotero_bibliographies" in policy and len(by_kind.get("ZOTERO_BIBL", [])) != policy["zotero_bibliographies"]:
            self.error("ZOTERO_BIBLIOGRAPHY_COUNT", actual=len(by_kind.get("ZOTERO_BIBL", [])))

    def footnotes(self):
        policy = self.spec.get("footnotes")
        if policy is None:
            return
        notes_root = self.parts.get("word/footnotes.xml")
        notes = [n for n in notes_root if n.tag == qn("w:footnote") and n.get(qn("w:type"), "normal") == "normal"] if notes_root is not None else []
        refs = [n for root in self.parts.values() for n in root.iter(qn("w:footnoteReference")) if visible(n)]
        ids = [n.get(qn("w:id")) for n in notes]
        reference_ids = [n.get(qn("w:id")) for n in refs]
        if Counter(ids) != Counter(reference_ids) or any(v != 1 for v in Counter(ids).values()):
            self.error("FOOTNOTE_REFERENCE_OR_ORPHAN", notes=ids, references=reference_ids)
        if len(refs) < policy.get("minimum_count", 0):
            self.error("FOOTNOTE_COUNT", actual=len(refs))
        defaults = {"pos": "pageBottom", "numRestart": "continuous", "numFmt": "decimal"}

        def overlay(values, pr):
            result = dict(values)
            if pr is not None:
                for key in values:
                    n = pr.find("w:" + key, NS)
                    if n is not None:
                        result[key] = n.get(qn("w:val"))
            return result

        defaults = overlay(defaults, self.document.settings.element.find("w:footnotePr", NS))
        effective = []
        for i, (_, sect) in enumerate(self.sections, 1):
            values = overlay(defaults, sect.find("w:footnotePr", NS))
            effective.append({"section": i, **values})
            for expected, key in (("position", "pos"), ("restart", "numRestart")):
                if expected in policy and values[key] != policy[expected]:
                    self.error("FOOTNOTE_SECTION_POLICY", section=i, property=key, actual=values[key], expected=policy[expected])
        for ref in refs:
            run = next((n for n in ref.iterancestors() if n.tag == qn("w:r")), None)
            para = self.paragraph(ref)
            if policy.get("superscript") and run is not None and para is not None:
                if effective_run(self.document, para, run).get("vertAlign") != "superscript":
                    self.error("FOOTNOTE_REFERENCE_NOT_SUPERSCRIPT", id=ref.get(qn("w:id")))
        self.observations["footnotes"] = {"count": len(refs), "effective_sections": effective}

    def captions(self):
        seen, pairs, sequence = set(), [], {}
        for rule in self.spec.get("captions", []):
            caption = self.anchor(rule["caption"])
            obj = self.anchor(rule["object"])
            if caption is None or obj is None:
                continue
            cp, op = self.paragraph(caption), self.top(obj)
            if cp is None:
                self.error("CAPTION_NOT_PARAGRAPH", id=rule["id"])
                continue
            label = normalized(f"{rule['label']}{rule['chapter']}-{rule['number']}")
            if label in seen:
                self.error("CAPTION_DUPLICATE_NUMBER", id=rule["id"])
            seen.add(label)
            display = normalized(text(cp))
            if not display.startswith(label) or (len(display) > len(label) and display[len(label)].isdigit()):
                self.error("CAPTION_DISPLAY_NUMBER", id=rule["id"], expected=label, actual=display)
            kind = rule["kind"]
            if op is None or (kind == "table" and op.tag != qn("w:tbl")) or (kind == "figure" and len(xpath(op, ".//w:drawing|.//w:pict")) != 1):
                self.error("CAPTION_OBJECT_KIND", id=rule["id"])
            ci, oi = self.index(caption), self.index(obj)
            if ci >= oi if kind == "table" else ci <= oi:
                self.error("CAPTION_POSITION", id=rule["id"], caption_index=ci, object_index=oi)
            fields = [f for f in self.fields if self.paragraph(f["node"]) is cp]
            seqs = [f for f in fields if re.match(r'\s*SEQ\s+(?:"' + re.escape(rule["label"]) + r'"|' + re.escape(rule["label"]) + r')(?:\s|$)', f["instruction"], re.I)]
            if len(seqs) != 1:
                self.error("CAPTION_SEQ", id=rule["id"], count=len(seqs))
            if rule.get("require_reference", True):
                name = rule["caption"].get("bookmark")
                reference = re.compile(r'\s*(?:REF|PAGEREF)\s+(?:"' + re.escape(name or "") + r'"|' + re.escape(name or "") + r')(?:\s|$)', re.I)
                if not name or not any(reference.match(f["instruction"]) for f in self.fields):
                    self.error("CAPTION_UNREFERENCED", id=rule["id"])
            if rule.get("keep_pair", True):
                first = cp if kind == "table" else self.paragraph(obj)
                if first is None or not paragraph_toggle(effective_paragraph(self.document, first), "keepNext"):
                    self.error("CAPTION_PAIR_NOT_KEPT", id=rule["id"])
                if kind == "figure" and paragraph_toggle(effective_paragraph(self.document, cp), "keepNext"):
                    self.review("FIGURE_CAPTION_DRAGS_FOLLOWING_TEXT", id=rule["id"])
                for between in self.children[min(ci, oi) + 1:max(ci, oi)]:
                    if (text(between).strip() or between.tag == qn("w:tbl")
                            or xpath(between, ".//w:drawing|.//w:pict")
                            or not paragraph_toggle(effective_paragraph(self.document, between), "keepNext")):
                        self.error("CAPTION_PAIR_INTERRUPTED", id=rule["id"], body_index=self.index(between))
                        break
            sequence.setdefault((rule["label"], rule["chapter"]), []).append((oi, rule["number"]))
            pairs.append({"id": rule["id"], "object_index": oi, "caption_index": ci})
        for group, values in sequence.items():
            actual = [v for _, v in sorted(values)]
            if actual != sorted(set(actual)):
                self.error("CAPTION_PHYSICAL_ORDER", group=list(group), numbers=actual)
            if self.spec.get("caption_sequence_complete") and actual != list(range(1, len(actual) + 1)):
                self.error("CAPTION_NUMBER_GAP", group=list(group), numbers=actual)
        self.observations["caption_pairs"] = pairs

    def bibliography(self):
        rule = self.spec.get("bibliography")
        if rule is None:
            return
        heading = self.anchor(rule["heading"])
        fields = [f for f in self.fields if self.field_kind(f) == "ZOTERO_BIBL"]
        if len(fields) != 1:
            self.error("BIBLIOGRAPHY_REQUIRES_ONE_LIVE_FIELD", actual=len(fields))
            return
        field = fields[0]
        if heading is not None and (field["part"] != "word/document.xml" or self.index(field["node"]) < self.index(heading)):
            self.error("BIBLIOGRAPHY_LOCATION")
        groups = {name: i for i, name in enumerate(rule["group_order"])}
        entries = rule["entries"]
        for entry in entries:
            if entry["language"] not in groups:
                self.error("BIBLIOGRAPHY_LANGUAGE_UNDECLARED", id=entry["id"])
            for key in ("author_display", "year", "title"):
                if normalized(entry[key]) not in normalized(entry["text"]):
                    self.error("BIBLIOGRAPHY_ENTRY_METADATA", id=entry["id"], property=key)
        ordered = sorted(entries, key=lambda e: (groups.get(e["language"], 999),
                         e["author_sort"].casefold(), e["year"].casefold(), e["title_sort"].casefold(), e["id"]))
        expected = normalized("".join(e["text"] for e in ordered))
        if expected != normalized(field["result"]):
            self.error("BIBLIOGRAPHY_VISIBLE_ORDER_OR_CONTENT", expected_ids=[e["id"] for e in ordered])
        if rule.get("entries_separate_paragraphs"):
            paragraphs = list(dict.fromkeys(self.paragraph(n) for n in field["result_nodes"]
                                            if n.tag == qn("w:t") and (n.text or "").strip()))
            actual = [normalized(text(p)) for p in paragraphs if p is not None]
            if actual != [normalized(e["text"]) for e in ordered]:
                self.error("BIBLIOGRAPHY_ENTRY_PARAGRAPHS", actual_paragraphs=len(actual), expected_entries=len(ordered))
        self.observations["bibliography"] = {"entries": len(entries), "groups": dict(Counter(e["language"] for e in entries)),
                                            "ordered_ids": [e["id"] for e in ordered],
                                            "sort_key_provenance": "declared in contract; linguistic/metadata correctness requires author review"}

    def objects(self):
        policy = self.spec.get("objects")
        if policy is None:
            return
        if policy.get("inline_images"):
            for drawing in self.body.iter(qn("w:drawing")):
                if not visible(drawing):
                    continue
                p = self.paragraph(drawing)
                if drawing.find("wp:anchor", NS) is not None:
                    self.error("IMAGE_FLOATING", body_index=self.index(drawing))
                if p is not None and effective_paragraph(self.document, p).get("spacing", {}).get("lineRule") == "exact":
                    self.error("IMAGE_EXACT_LINE_HEIGHT", body_index=self.index(drawing))
                number = self.section_number(drawing)
                extent = drawing.find(".//wp:extent", NS)
                if number and extent is not None:
                    section = self.document.sections[number - 1]
                    width = int(section.page_width - section.left_margin - section.right_margin)
                    height = int(section.page_height - section.top_margin - section.bottom_margin)
                    if int(extent.get("cx", "0")) > width or int(extent.get("cy", "0")) > height:
                        self.error("IMAGE_EXCEEDS_SECTION_AREA", body_index=self.index(drawing))
        for rule in policy.get("tables", []):
            anchor = self.anchor(rule["anchor"])
            table = self.top(anchor) if anchor is not None else None
            if table is None or table.tag != qn("w:tbl"):
                self.error("TABLE_ANCHOR_NOT_TABLE", anchor=rule["anchor"])
                continue
            for row in table.findall("w:tr", NS):
                height = row.find("w:trPr/w:trHeight", NS)
                if rule.get("growing_rows", True) and height is not None and height.get(qn("w:hRule")) == "exact":
                    self.error("TABLE_EXACT_ROW_HEIGHT", anchor=rule["anchor"])
            rows = table.findall("w:tr", NS)
            if rule.get("repeat_header") and rows:
                header = rows[0].find("w:trPr/w:tblHeader", NS)
                if header is None or not on(header.get(qn("w:val"), "true")):
                    self.error("TABLE_HEADER_NOT_REPEATED", anchor=rule["anchor"])
            for cell in table.iter(qn("w:tc")):
                content = text(cell)
                align = "left" if HAN.search(content) else "right"
                for p in cell.findall("w:p", NS):
                    props = effective_paragraph(self.document, p)
                    if rule.get("zero_indents", True):
                        for name in ("firstLine", "hanging", "firstLineChars", "hangingChars"):
                            if float(props.get("ind", {}).get(name, "0")) != 0:
                                self.error("TABLE_INHERITED_INDENT", anchor=rule["anchor"], property=name)
                    if rule.get("han_left_nonhan_right") and content.strip() and props.get("jc", {}).get("val", "left") != align:
                        self.error("TABLE_ALIGNMENT", anchor=rule["anchor"], expected=align, text=content)
            if rule.get("plain_header") and rows:
                for p in rows[0].iter(qn("w:p")):
                    for r in p.iter(qn("w:r")):
                        if text(r).strip() and effective_run(self.document, p, r).get("b"):
                            self.error("TABLE_HEADER_BOLD", anchor=rule["anchor"])
            # Conditional table styles can affect Word's rendering. Direct and
            # paragraph-style checks alone cannot certify that extra cascade.
            style = table.find("w:tblPr/w:tblStyle", NS)
            if style is not None:
                sid = style.get(qn("w:val"))
                candidate = next((s.element for s in self.document.styles if s.style_id == sid), None)
                if candidate is not None and candidate.find("w:tblStylePr", NS) is not None:
                    self.review("CONDITIONAL_TABLE_STYLE_REQUIRES_RENDER_CHECK", style=sid, anchor=rule["anchor"])

    def run(self):
        before = digest(self.path)
        self.scan()
        self.blocks()
        self.fields_check()
        self.footnotes()
        self.captions()
        self.bibliography()
        self.objects()
        if digest(self.path) != before:
            self.error("INPUT_CHANGED_DURING_AUDIT")
        return {"schema_version": "1.0", "mode": "read_only_structural", "docx_sha256": before,
                "contract_sha256": hashlib.sha256(json.dumps(self.spec, ensure_ascii=False, sort_keys=True).encode()).hexdigest(),
                "all_pass": not self.errors, "status": "FAIL" if self.errors else "REVIEW" if self.reviews else "PASS",
                "errors": self.errors, "review": self.reviews, "observations": self.observations,
                "not_checked": ["native pagination and TOC page accuracy", "visual aesthetics and first-reference logic",
                                "citation support and bibliographic metadata provenance", "linguistic correctness of pinyin sort keys"],
                "omitted_contract_areas": [k for k in ("footnotes", "captions", "bibliography", "objects", "sections") if k not in self.spec]}


def audit(path, spec):
    return Audit(path, spec).run()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--docx", required=True)
    parser.add_argument("--contract", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    result = audit(args.docx, json.loads(Path(args.contract).read_text(encoding="utf-8")))
    output = Path(args.output)
    # No accidental replacement of a previous report or the input.
    with output.open("x", encoding="utf-8") as handle:
        handle.write(json.dumps(result, ensure_ascii=False, indent=2) + "\n")
    print(json.dumps({"status": result["status"], "output": str(output.resolve())}))
    return 0 if result["all_pass"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
