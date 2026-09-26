#!/usr/bin/env python3
"""Check rendered glyph fonts and header/footer geometry in the final PDF."""
from __future__ import annotations
import argparse
import hashlib
import json
from pathlib import Path
import re


def normalized_font(name):
    return re.sub(r"[^a-z0-9]", "", re.sub(r"^[A-Z]{6}\+", "", name).lower())


def check_page(chars, spec):
    errors = []
    measurements = []
    if not spec.get("font_regions") and not spec.get("stories"):
        errors.append({"type": "NO_RULES_IN_PAGE_SCOPE"})
    for rule in spec.get("font_regions", []):
        left, top, right, bottom = rule["box"]
        selected = [c for c in chars if left <= c["x0"] < right and top <= c["top"] < bottom and c["text"].strip()]
        if not selected:
            errors.append({"type": "NO_TEXT_IN_FONT_SCOPE", "rule": rule.get("id")})
        checked = 0
        for c in selected:
            script = "han" if re.search(r"[\u3400-\u9fff\U00020000-\U0002ffff]", c["text"]) else "latin" if re.search(r"[A-Za-z0-9]", c["text"]) else None
            if script and rule.get(script):
                checked += 1
                if normalized_font(c["fontname"]) not in {normalized_font(f) for f in rule[script]}:
                    errors.append({"type": "FONT_MISMATCH", "text": c["text"], "actual": c["fontname"], "expected": rule[script], "x": c["x0"], "y": c["top"]})
        measurements.append({"id": rule.get("id"), "glyphs_checked": checked})
        if checked == 0:
            errors.append({"type": "NO_MATCHING_SCRIPT_GLYPHS", "rule": rule.get("id")})
    for rule in spec.get("stories", []):
        left, top, right, bottom = rule["box"]
        region = sorted((c for c in chars if left <= c["x0"] < right and top <= c["top"] < bottom and c["text"].strip()), key=lambda c: (round(c["top"], 1), c["x0"]))
        target = "".join(rule["text"].split())
        actual = "".join(c["text"] for c in region)
        start = actual.find(target)
        if start < 0 or not target:
            errors.append({"type": "STORY_TEXT_NOT_FOUND", "text": rule["text"]})
            continue
        # PDF 字符对象通常一字一项；按文本长度映射，兼容多字符字形。
        selected, offset = [], 0
        for c in region:
            end = offset + len(c["text"])
            if end > start and offset < start + len(target):
                selected.append(c)
            offset = end
        x0, x1 = min(c["x0"] for c in selected), max(c["x1"] for c in selected)
        alignment = rule["alignment"]
        if alignment not in {"center", "left", "right"}:
            raise ValueError("Unsupported story alignment")
        area = rule["text_area"]
        expected = (area[0] + area[1]) / 2 if alignment == "center" else area[0] if alignment == "left" else area[1]
        observed = (x0 + x1) / 2 if alignment == "center" else x0 if alignment == "left" else x1
        delta = observed - expected
        measurements.append({"id": rule.get("id"), "offset_pt": delta, "expected": expected, "observed": observed})
        if abs(delta) > rule.get("tolerance_pt", 1):
            errors.append({"type": "STORY_GEOMETRY_MISMATCH", "id": rule.get("id"), "offset_pt": delta})
    return {"pass": not errors, "errors": errors, "measurements": measurements}


def audit(pdf, spec):
    import pdfplumber
    pages = []
    with pdfplumber.open(pdf) as document:
        for rule in spec["pages"]:
            number = rule["page"]
            if not 1 <= number <= len(document.pages):
                pages.append({"page": number, "pass": False, "errors": [{"type": "PAGE_MISSING"}]})
                continue
            page = document.pages[number - 1]
            result = check_page(page.chars, rule)
            pages.append({"page": number, "width": float(page.width), "height": float(page.height), **result})
    return {"mode": "real", "status": "PASS" if pages and all(p["pass"] for p in pages) else "FAIL",
            "pdf_sha256": hashlib.sha256(Path(pdf).read_bytes()).hexdigest(), "pages": pages,
            "coverage": "explicit glyph regions and story rules; other pages/objects require separate visual review"}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--pdf", required=True)
    parser.add_argument("--spec", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    result = audit(args.pdf, json.loads(Path(args.spec).read_text()))
    Path(args.output).write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n")
    print(json.dumps({"status": result["status"], "output": str(Path(args.output).resolve())}))
    return 0 if result["status"] == "PASS" else 2


if __name__ == "__main__":
    raise SystemExit(main())
