#!/usr/bin/env python3
"""Post-processing checks and fixes for MathType equations in DOCX files.

Inline MathType objects are usually taller than the text line. In a paragraph with *exact* line
spacing (CJGE body text uses exactly 15.6 pt) Word clips everything above the line height, so the
top of fractions, radicals and superscripts disappears. ``fix`` switches exactly those paragraphs
(including table-cell paragraphs) to "at least" spacing with the same value; ``check`` only reports
them. The spacing may come from the paragraph itself, from its style chain or from docDefaults.

Usage: docx_postprocess.py fix INPUT OUTPUT [--overwrite]
       docx_postprocess.py check DOCUMENT
Prints one JSON result line.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import tempfile
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from mathtype_prefs import NS, _css_length_pt  # noqa: E402

W = "{%s}" % NS["w"]
BODY_PARTS = r"word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml"


def _spacing_from(ppr) -> dict:
    if ppr is None:
        return {}
    spacing = ppr.find("w:spacing", NS)
    if spacing is None:
        return {}
    return {key: spacing.get(W + key) for key in ("line", "lineRule") if spacing.get(W + key) is not None}


class _StyleSpacing:
    """Resolve the inherited line spacing of a paragraph style (basedOn chain, then docDefaults)."""

    def __init__(self, styles_xml: bytes | None) -> None:
        from lxml import etree

        self.styles: dict[str, object] = {}
        self.default_style = None
        self.defaults: dict = {}
        if not styles_xml:
            return
        root = etree.fromstring(styles_xml)
        self.defaults = _spacing_from(root.find("w:docDefaults/w:pPrDefault/w:pPr", NS))
        for style in root.findall("w:style", NS):
            if style.get(W + "type") != "paragraph":
                continue
            self.styles[style.get(W + "styleId")] = style
            if style.get(W + "default") in ("1", "true"):
                self.default_style = style.get(W + "styleId")

    def resolve(self, style_id: str | None) -> dict:
        result: dict = {}
        seen: set[str] = set()
        current = style_id or self.default_style
        while current and current not in seen and current in self.styles:
            seen.add(current)
            style = self.styles[current]
            for key, value in _spacing_from(style.find("w:pPr", NS)).items():
                result.setdefault(key, value)
            based = style.find("w:basedOn", NS)
            current = based.get(W + "val") if based is not None else None
        for key, value in self.defaults.items():
            result.setdefault(key, value)
        return result


def _equation_heights(paragraph) -> list[float]:
    heights = []
    for obj in paragraph.iter(W + "object"):
        ole = obj.find("o:OLEObject", NS)
        shape = obj.find("v:shape", NS)
        if ole is None or shape is None or not str(ole.get("ProgID", "")).startswith("Equation.DSMT"):
            continue
        height = _css_length_pt(shape.get("style", ""), "height")
        if height is not None:
            heights.append(height)
    return heights


def _clipped_paragraphs(root, styles: _StyleSpacing) -> list[tuple[object, float, float]]:
    """Paragraphs with exact line spacing lower than the tallest MathType object they contain."""
    found = []
    for paragraph in root.iter(W + "p"):
        heights = _equation_heights(paragraph)
        if not heights:
            continue
        ppr = paragraph.find("w:pPr", NS)
        style = ppr.find("w:pStyle", NS) if ppr is not None else None
        spacing = styles.resolve(style.get(W + "val") if style is not None else None)
        spacing.update(_spacing_from(ppr))
        if spacing.get("lineRule") != "exact" or not spacing.get("line"):
            continue
        line_pt = int(spacing["line"]) / 20.0
        tallest = max(heights)
        if tallest > line_pt + 0.5:
            found.append((paragraph, line_pt, tallest))
    return found


def _iter_parts(package: zipfile.ZipFile):
    for name in package.namelist():
        if re.fullmatch(BODY_PARTS, name):
            yield name


def check(document_path: str) -> dict:
    from lxml import etree

    path = Path(document_path).resolve()
    with zipfile.ZipFile(path) as package:
        styles = _StyleSpacing(package.read("word/styles.xml") if "word/styles.xml" in package.namelist() else None)
        clipped = []
        for part in _iter_parts(package):
            root = etree.fromstring(package.read(part))
            for paragraph, line_pt, tallest in _clipped_paragraphs(root, styles):
                text = "".join(t.text or "" for t in paragraph.iter(W + "t"))[:40]
                clipped.append({"part": part, "line_pt": line_pt, "tallest_equation_pt": round(tallest, 2), "text": text})
    warnings = []
    if clipped:
        warnings.append(
            f"{len(clipped)} paragraph(s) use exact line spacing lower than an inline MathType equation, so the "
            "equation is clipped; run fix_mathtype_line_spacing (render does this automatically)."
        )
    return {"ok": True, "action": "check-line-spacing", "document_path": str(path),
            "clipped_paragraphs": len(clipped), "examples": clipped[:5], "warnings": warnings}


def fix(input_path: str, output_path: str, overwrite: bool = False) -> dict:
    from lxml import etree

    source = Path(input_path).resolve()
    destination = Path(output_path).resolve()
    if destination.exists() and destination != source and not overwrite:
        raise FileExistsError(f"Output exists (pass --overwrite): {destination}")
    if destination == source and not overwrite:
        raise FileExistsError("Fixing in place requires --overwrite.")
    fixed = 0
    with zipfile.ZipFile(source) as package:
        styles = _StyleSpacing(package.read("word/styles.xml") if "word/styles.xml" in package.namelist() else None)
        replaced: dict[str, bytes] = {}
        for part in _iter_parts(package):
            root = etree.fromstring(package.read(part))
            hits = _clipped_paragraphs(root, styles)
            for paragraph, line_pt, _tallest in hits:
                ppr = paragraph.find("w:pPr", NS)
                if ppr is None:
                    ppr = etree.Element(W + "pPr")
                    paragraph.insert(0, ppr)
                spacing = ppr.find("w:spacing", NS)
                if spacing is None:
                    spacing = etree.SubElement(ppr, W + "spacing")
                spacing.set(W + "line", str(round(line_pt * 20)))
                spacing.set(W + "lineRule", "atLeast")
                fixed += 1
            if hits:
                replaced[part] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        handle, temporary = tempfile.mkstemp(suffix=".docx", dir=destination.parent)
        os.close(handle)
        try:
            with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as out:
                for item in package.infolist():
                    out.writestr(item, replaced.get(item.filename) or package.read(item.filename))
        except BaseException:
            os.unlink(temporary)
            raise
    os.replace(temporary, destination)
    return {"ok": True, "action": "fix-line-spacing", "input_path": str(source), "output_path": str(destination),
            "paragraphs_fixed": fixed}



def _utf8_stdio() -> None:
    """Windows consoles default to an ANSI code page; JSON results may contain CJK paths."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass


def main() -> int:
    _utf8_stdio()
    parser = argparse.ArgumentParser()
    sub = parser.add_subparsers(dest="cmd", required=True)
    f = sub.add_parser("fix")
    f.add_argument("input")
    f.add_argument("output")
    f.add_argument("--overwrite", action="store_true")
    c = sub.add_parser("check")
    c.add_argument("document")
    args = parser.parse_args()
    try:
        result = fix(args.input, args.output, args.overwrite) if args.cmd == "fix" else check(args.document)
    except Exception as exc:
        result = {"ok": False, "action": f"{args.cmd}-line-spacing", "error": f"{type(exc).__name__}: {exc}"}
    sys.stdout.write(json.dumps(result, ensure_ascii=False) + "\n")
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
