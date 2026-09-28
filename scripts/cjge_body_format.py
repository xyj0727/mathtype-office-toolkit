#!/usr/bin/env python3
"""CJGE body-text formatting and a compact formatting report for DOCX files.

``apply``  formats everything except equations from a body profile (default
           config/cjge_body_profile.json): page size and margins, fonts, sizes, line spacing,
           indents and alignment per paragraph role, and three-line tables. Runs that hold MathType
           objects and 1x3 equation tables are never touched; paragraphs holding inline equations get
           "at least" spacing so the equations are not clipped.
``report`` summarises the effective formatting per paragraph role (title, heading levels, body,
           list, caption, table, reference) instead of dumping raw XML, and with a profile lists the
           deviations from it.

Paragraph roles: heading levels come from the style definitions (heading N / outline level);
the title is the first non-empty paragraph before any heading; "front" is the short block right
after the title; captions start with 图/表/Fig./Figure/Table and a number; the reference list is the
part after a 参考文献/References heading.

Usage: cjge_body_format.py apply INPUT OUTPUT [--profile P.json] [--overwrite]
       cjge_body_format.py report DOCUMENT [--profile P.json|cjge]
Prints one JSON result line.
"""
from __future__ import annotations

import argparse
import copy
import json
import os
import re
import sys
import tempfile
import zipfile
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from mathtype_scan import REFERENCE_HEADING, _heading_levels  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_PROFILE = ROOT / "config" / "cjge_body_profile.json"
WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % WNS

# CT_RPr / CT_PPr / CT_TblPr / CT_TcPr child order (ECMA-376); Word rejects out-of-order children.
RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline",
             "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w",
             "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign",
             "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"]
PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
             "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
             "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid",
             "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
             "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"]
TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize",
               "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook"]
TCPR_ORDER = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
              "textDirection", "tcFitText", "vAlign", "hideMark"]
JC = {"left": "left", "center": "center", "right": "right", "justify": "both"}
CAPTION = re.compile(r"^\s*(图|表|Fig\.?|Figure|Table)\s*\d")


def _child(parent, name: str, order: list[str], create: bool = True):
    from lxml import etree

    found = parent.find(W + name)
    if found is not None or not create:
        return found
    element = etree.Element(W + name)
    rank = order.index(name)
    for position, existing in enumerate(parent):
        local = existing.tag.split("}")[-1]
        if local in order and order.index(local) > rank:
            parent.insert(position, element)
            return element
    parent.append(element)
    return element


def _remove(parent, name: str) -> None:
    for element in parent.findall(W + name):
        parent.remove(element)


def _ensure(parent, name: str, order: list[str], position: int = 0):
    from lxml import etree

    found = parent.find(W + name)
    if found is None:
        found = etree.Element(W + name)
        parent.insert(position, found)
    return found


# ---------------------------------------------------------------- document structure

def _is_equation_table(table) -> bool:
    rows = table.findall(W + "tr")
    if not rows:
        return False
    for row in rows:
        cells = row.findall(W + "tc")
        if len(cells) != 3:
            return False
        if not any(_has_equation(p) for p in cells[1].iter(W + "p")):
            return False
    return True


def _has_equation(element) -> bool:
    for ole in element.iter("{urn:schemas-microsoft-com:office:office}OLEObject"):
        if str(ole.get("ProgID", "")).startswith("Equation.DSMT"):
            return True
    return False


def _text(paragraph) -> str:
    return "".join(t.text or "" for t in paragraph.iter(W + "t"))


def style_names(styles_xml: bytes | None) -> dict[str, str]:
    from lxml import etree

    if not styles_xml:
        return {}
    names = {}
    for style in etree.fromstring(styles_xml).findall(W + "style"):
        name = style.find(W + "name")
        names[style.get(W + "styleId")] = (name.get(W + "val") if name is not None else "") or ""
    return names


def classify(root, levels: dict[str, int], names: dict[str, str] | None = None) -> list[tuple[object, str]]:
    """(paragraph, role) for every paragraph; role 'skip' for equation-table paragraphs."""
    body = root.find(W + "body")
    roles: list[tuple[object, str]] = []
    equation_tables = {id(t) for t in body.iter(W + "tbl") if _is_equation_table(t)}
    seen_heading = False
    title_done = False
    front_open = False
    reference_level = None
    for paragraph in body.iter(W + "p"):
        tables = [a for a in paragraph.iterancestors() if a.tag == W + "tbl"]
        if any(id(t) in equation_tables for t in tables):
            roles.append((paragraph, "skip"))
            continue
        style = paragraph.find(f"{W}pPr/{W}pStyle")
        style_id = style.get(W + "val") if style is not None else None
        level = levels.get(style_id)
        text = _text(paragraph).strip()
        if tables:
            row = [a for a in paragraph.iterancestors() if a.tag == W + "tr"][0]
            first_row = tables[0].find(W + "tr") is row
            roles.append((paragraph, "table_header" if first_row else "table"))
            continue
        if level is not None:
            seen_heading = True
            front_open = False
            if reference_level is not None and level <= reference_level:
                reference_level = None
            if REFERENCE_HEADING.match(text):
                reference_level = level
            roles.append((paragraph, f"heading{min(level, 3)}"))
            continue
        if reference_level is not None:
            roles.append((paragraph, "reference"))
            continue
        if not title_done and not seen_heading and text:
            title_done = True
            front_open = True
            roles.append((paragraph, "title"))
            continue
        if front_open and text and len(text) <= 120 and not text.startswith("【"):
            roles.append((paragraph, "front"))
            continue
        front_open = False
        if text and CAPTION.match(text) and len(text) <= 120:
            roles.append((paragraph, "caption"))
            continue
        ppr = paragraph.find(W + "pPr")
        numbered = ppr is not None and ppr.find(W + "numPr") is not None
        style_name = (names or {}).get(style_id, "") if style_id else ""
        style_is_list = bool(style_id and re.search(r"(?i)list|列表", f"{style_id} {style_name}"))
        roles.append((paragraph, "list" if numbered or style_is_list else "body"))
    return roles


# ---------------------------------------------------------------- apply

def _format_run(run, spec: dict, remove_color: bool) -> None:
    rpr = _ensure(run, "rPr", RPR_ORDER)
    fonts = _child(rpr, "rFonts", RPR_ORDER)
    for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        fonts.attrib.pop(W + attr, None)
    for attr in ("ascii", "hAnsi", "cs"):
        fonts.set(W + attr, spec["latin"])
    fonts.set(W + "eastAsia", spec["east_asia"])
    half_points = str(round(float(spec["size_pt"]) * 2))
    _child(rpr, "sz", RPR_ORDER).set(W + "val", half_points)
    _child(rpr, "szCs", RPR_ORDER).set(W + "val", half_points)
    if spec.get("bold") is not None:
        for name in ("b", "bCs"):
            _remove(rpr, name)
            element = _child(rpr, name, RPR_ORDER)
            if not spec["bold"]:  # explicit off: heading styles are often bold themselves
                element.set(W + "val", "0")
    if remove_color:
        _remove(rpr, "color")


def _format_paragraph(paragraph, spec: dict) -> None:
    ppr = _ensure(paragraph, "pPr", PPR_ORDER)
    spacing = _child(ppr, "spacing", PPR_ORDER)
    for attr in ("beforeLines", "afterLines", "beforeAutospacing", "afterAutospacing"):
        spacing.attrib.pop(W + attr, None)
    spacing.set(W + "before", str(round(float(spec.get("space_before_pt", 0)) * 20)))
    spacing.set(W + "after", str(round(float(spec.get("space_after_pt", 0)) * 20)))
    if spec.get("line_rule"):
        rule = spec["line_rule"]
        spacing.set(W + "lineRule", {"exact": "exact", "at_least": "atLeast", "auto": "auto"}[rule])
        spacing.set(W + "line", str(round(float(spec["line_pt"]) * (240 / 12 if rule == "auto" else 20))))
    ind = _child(ppr, "ind", PPR_ORDER)
    for attr in ("firstLine", "firstLineChars", "hanging", "hangingChars"):
        ind.attrib.pop(W + attr, None)
    if spec.get("hanging_pt"):
        ind.set(W + "left", str(round(float(spec["hanging_pt"]) * 20)))
        ind.set(W + "hanging", str(round(float(spec["hanging_pt"]) * 20)))
    elif spec.get("first_line_chars"):
        chars = float(spec["first_line_chars"])
        ind.set(W + "firstLineChars", str(round(chars * 100)))
        ind.set(W + "firstLine", str(round(chars * float(spec["size_pt"]) * 20)))
    if spec.get("align"):
        _child(ppr, "jc", PPR_ORDER).set(W + "val", JC[spec["align"]])


def _three_line(table, spec: dict) -> None:
    tblpr = _ensure(table, "tblPr", TBLPR_ORDER)
    _remove(tblpr, "tblStyle")
    if spec.get("align"):
        _child(tblpr, "jc", TBLPR_ORDER).set(W + "val", JC[spec["align"]])
    _remove(tblpr, "tblBorders")
    borders = _child(tblpr, "tblBorders", TBLPR_ORDER)
    outer = str(round(float(spec.get("outer_rule_pt", 1.5)) * 8))
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = _child(borders, side, ["top", "left", "start", "bottom", "right", "end", "insideH", "insideV"])
        if side in ("top", "bottom"):
            element.set(W + "val", "single")
            element.set(W + "sz", outer)
            element.set(W + "space", "0")
            element.set(W + "color", "000000")
        else:
            element.set(W + "val", "nil")
    rows = table.findall(W + "tr")
    header = str(round(float(spec.get("header_rule_pt", 0.75)) * 8))
    for index, row in enumerate(rows):
        for cell in row.findall(W + "tc"):
            tcpr = _ensure(cell, "tcPr", TCPR_ORDER)
            _remove(tcpr, "tcBorders")
            _remove(tcpr, "shd")
            if index == 0 and len(rows) > 1:
                cell_borders = _child(tcpr, "tcBorders", TCPR_ORDER)
                bottom = _child(cell_borders, "bottom", ["top", "left", "start", "bottom", "right", "end"])
                bottom.set(W + "val", "single")
                bottom.set(W + "sz", header)
                bottom.set(W + "space", "0")
                bottom.set(W + "color", "000000")


def _set_page(root, page: dict) -> None:
    for sect in root.iter(W + "sectPr"):
        size = sect.find(W + "pgSz")
        if size is not None:
            landscape = size.get(W + "orient") == "landscape"
            width, height = round(page["width_cm"] * 567), round(page["height_cm"] * 567)
            size.set(W + "w", str(height if landscape else width))
            size.set(W + "h", str(width if landscape else height))
        margin = sect.find(W + "pgMar")
        if margin is not None:
            for side in ("top", "bottom", "left", "right"):
                margin.set(W + side, str(round(page[f"{side}_cm"] * 567)))


def _set_default_fonts(styles_root, spec: dict) -> None:
    rpr_default = styles_root.find(f"{W}docDefaults/{W}rPrDefault/{W}rPr")
    if rpr_default is None:
        return
    fonts = _child(rpr_default, "rFonts", RPR_ORDER)
    for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        fonts.attrib.pop(W + attr, None)
    for attr in ("ascii", "hAnsi", "cs"):
        fonts.set(W + attr, spec["latin"])
    fonts.set(W + "eastAsia", spec["east_asia"])


def load_profile(path: str | None) -> dict:
    source = DEFAULT_PROFILE if not path or path.lower() == "cjge" else Path(path)
    return json.loads(Path(source).read_text(encoding="utf-8-sig"))


def apply(input_path: str, output_path: str, profile_path: str | None = None, overwrite: bool = False) -> dict:
    from lxml import etree

    import docx_postprocess

    source = Path(input_path).resolve()
    destination = Path(output_path).resolve()
    if destination == source:
        raise ValueError("Write the formatted DOCX to a new path; the source is preserved.")
    if destination.exists() and not overwrite:
        raise FileExistsError(f"Output exists (pass overwrite): {destination}")
    profile = load_profile(profile_path)
    remove_color = bool(profile.get("remove_text_color", True))
    with zipfile.ZipFile(source) as package:
        names = package.namelist()
        root = etree.fromstring(package.read("word/document.xml"))
        styles_xml = package.read("word/styles.xml") if "word/styles.xml" in names else None
        levels = _heading_levels(styles_xml)
        counts: Counter = Counter()
        skipped_runs = 0
        for paragraph, role in classify(root, levels, style_names(styles_xml)):
            if role == "skip":
                counts["equation_table_paragraphs_skipped"] += 1
                continue
            spec = profile["roles"][role]
            _format_paragraph(paragraph, spec)
            for run in paragraph.iter(W + "r"):
                if _has_equation(run) or run.find(W + "object") is not None:
                    skipped_runs += 1
                    continue
                _format_run(run, spec, remove_color)
            counts[role] += 1
        tables = 0
        if profile.get("tables", {}).get("three_line"):
            for table in root.find(W + "body").iter(W + "tbl"):
                if not _is_equation_table(table):
                    _three_line(table, profile["tables"])
                    tables += 1
        _set_page(root, profile["page"])
        replaced = {"word/document.xml": etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)}
        if styles_xml:
            styles_root = etree.fromstring(styles_xml)
            _set_default_fonts(styles_root, profile["roles"]["body"])
            replaced["word/styles.xml"] = etree.tostring(styles_root, xml_declaration=True, encoding="UTF-8",
                                                         standalone=True)
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
    spacing = docx_postprocess.fix(str(destination), str(destination), overwrite=True)
    return {"ok": True, "action": "apply-cjge-body-format", "input_path": str(source),
            "output_path": str(destination), "profile": profile.get("name"), "paragraphs": dict(counts),
            "three_line_tables": tables, "equation_runs_untouched": skipped_runs,
            "inline_equation_paragraphs_set_to_at_least": spacing.get("paragraphs_fixed", 0)}


# ---------------------------------------------------------------- report

class _Styles:
    """Effective run and paragraph properties through the style chain and docDefaults."""

    def __init__(self, styles_xml: bytes | None) -> None:
        from lxml import etree

        self.by_id: dict[str, object] = {}
        self.default_paragraph = None
        self.rpr_default = None
        self.ppr_default = None
        if not styles_xml:
            return
        root = etree.fromstring(styles_xml)
        self.rpr_default = root.find(f"{W}docDefaults/{W}rPrDefault/{W}rPr")
        self.ppr_default = root.find(f"{W}docDefaults/{W}pPrDefault/{W}pPr")
        for style in root.findall(W + "style"):
            self.by_id[style.get(W + "styleId")] = style
            if style.get(W + "type") == "paragraph" and style.get(W + "default") in ("1", "true"):
                self.default_paragraph = style.get(W + "styleId")

    def chain(self, style_id: str | None) -> list:
        chain, seen = [], set()
        current = style_id
        while current and current not in seen and current in self.by_id:
            seen.add(current)
            style = self.by_id[current]
            chain.append(style)
            based = style.find(W + "basedOn")
            current = based.get(W + "val") if based is not None else None
        return chain

    def run_props(self, run, paragraph) -> dict:
        layers = [run.find(W + "rPr")]
        rstyle = run.find(f"{W}rPr/{W}rStyle")
        if rstyle is not None:
            layers += [s.find(W + "rPr") for s in self.chain(rstyle.get(W + "val"))]
        pstyle = paragraph.find(f"{W}pPr/{W}pStyle")
        layers += [s.find(W + "rPr") for s in self.chain(pstyle.get(W + "val") if pstyle is not None
                                                          else self.default_paragraph)]
        layers.append(self.rpr_default)
        result: dict = {}
        for layer in layers:
            if layer is None:
                continue
            fonts = layer.find(W + "rFonts")
            if fonts is not None:
                for key, attrs in (("east_asia", ("eastAsia", "eastAsiaTheme")), ("latin", ("ascii", "asciiTheme"))):
                    if key not in result:
                        for attr in attrs:
                            if fonts.get(W + attr):
                                result[key] = fonts.get(W + attr) if attr in ("eastAsia", "ascii") \
                                    else "theme:" + fonts.get(W + attr)
                                break
            size = layer.find(W + "sz")
            if size is not None and "size_pt" not in result:
                result["size_pt"] = int(size.get(W + "val")) / 2
            bold = layer.find(W + "b")
            if bold is not None and "bold" not in result:
                result["bold"] = bold.get(W + "val") not in ("0", "false")
        result.setdefault("size_pt", 10.0)
        result.setdefault("bold", False)
        return result

    def paragraph_props(self, paragraph) -> dict:
        pstyle = paragraph.find(f"{W}pPr/{W}pStyle")
        layers = [paragraph.find(W + "pPr")]
        layers += [s.find(W + "pPr") for s in self.chain(pstyle.get(W + "val") if pstyle is not None
                                                          else self.default_paragraph)]
        layers.append(self.ppr_default)
        result: dict = {}
        for layer in layers:
            if layer is None:
                continue
            spacing = layer.find(W + "spacing")
            if spacing is not None:
                if "line_rule" not in result and spacing.get(W + "line"):
                    rule = spacing.get(W + "lineRule") or "auto"
                    value = int(spacing.get(W + "line"))
                    result["line_rule"] = {"atLeast": "at_least"}.get(rule, rule)
                    result["line"] = round(value / 240, 2) if rule == "auto" else round(value / 20, 1)
            ind = layer.find(W + "ind")
            if ind is not None and "first_line_chars" not in result:
                if ind.get(W + "firstLineChars"):
                    result["first_line_chars"] = int(ind.get(W + "firstLineChars")) / 100
                elif ind.get(W + "hanging"):
                    result["first_line_chars"] = 0
                    result["hanging_pt"] = int(ind.get(W + "hanging")) / 20
                elif ind.get(W + "firstLine"):
                    result["first_line_pt"] = int(ind.get(W + "firstLine")) / 20
            jc = layer.find(W + "jc")
            if jc is not None and "align" not in result:
                result["align"] = {"both": "justify", "start": "left", "end": "right"}.get(jc.get(W + "val"),
                                                                                         jc.get(W + "val"))
        result.setdefault("align", "left")
        return result


def _font_matches(actual: str | None, wanted: str) -> bool:
    aliases = {"宋体": {"宋体", "SimSun"}, "黑体": {"黑体", "SimHei"}, "楷体": {"楷体", "KaiTi", "楷体_GB2312"},
               "仿宋": {"仿宋", "FangSong"}}
    return actual in aliases.get(wanted, {wanted})


def report(document_path: str, profile_path: str | None = None, examples: int = 3) -> dict:
    from lxml import etree

    path = Path(document_path).resolve()
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        root = etree.fromstring(package.read("word/document.xml"))
        styles_xml = package.read("word/styles.xml") if "word/styles.xml" in names else None
    styles = _Styles(styles_xml)
    levels = _heading_levels(styles_xml)
    profile = load_profile(profile_path) if profile_path else None
    roles: dict[str, dict] = {}
    deviations: dict[str, dict] = {}
    for paragraph, role in classify(root, levels, style_names(styles_xml)):
        if role == "skip":
            continue
        text = _text(paragraph).strip()
        entry = roles.setdefault(role, {"paragraphs": 0, "runs": Counter(), "layout": Counter(), "examples": {}})
        entry["paragraphs"] += 1
        pprops = styles.paragraph_props(paragraph)
        layout_key = json.dumps(pprops, ensure_ascii=False, sort_keys=True)
        entry["layout"][layout_key] += 1
        entry["examples"].setdefault(layout_key, text[:30])
        spec = profile["roles"].get(role) if profile else None
        problems = []
        if spec:
            if spec.get("line_rule") and (pprops.get("line_rule"), pprops.get("line")) != (spec["line_rule"], spec["line_pt"]) \
                    and not (_has_equation(paragraph) and pprops.get("line_rule") == "at_least"):
                problems.append(f"line {pprops.get('line_rule')} {pprops.get('line')}")
            if spec.get("align") and pprops.get("align") != spec["align"]:
                problems.append(f"align {pprops.get('align')}")
        for run in paragraph.iter(W + "r"):
            if run.find(W + "object") is not None:
                continue
            content = "".join(t.text or "" for t in run.findall(W + "t"))
            if not content.strip():
                continue
            rprops = styles.run_props(run, paragraph)
            key = json.dumps(rprops, ensure_ascii=False, sort_keys=True)
            entry["runs"][key] += 1
            entry["examples"].setdefault(key, content[:30])
            if spec:
                has_cjk = re.search(r"[㐀-鿿]", content)
                has_latin = re.search(r"[A-Za-z0-9]", content)
                if has_cjk and not _font_matches(rprops.get("east_asia"), spec["east_asia"]):
                    problems.append(f"East Asian font {rprops.get('east_asia')}")
                if has_latin and not _font_matches(rprops.get("latin"), spec["latin"]):
                    problems.append(f"Latin font {rprops.get('latin')}")
                if abs(float(rprops["size_pt"]) - float(spec["size_pt"])) > 0.01 and not run.find(f"{W}rPr/{W}vertAlign") is not None:
                    problems.append(f"size {rprops['size_pt']:g} pt")
                if spec.get("bold") is not None and rprops.get("bold") != spec["bold"]:
                    problems.append(f"bold {rprops.get('bold')}")
        if problems:
            dev = deviations.setdefault(role, {"paragraphs": 0, "issues": Counter(), "examples": []})
            dev["paragraphs"] += 1
            for problem in set(problems):
                dev["issues"][problem] += 1
            if len(dev["examples"]) < examples:
                dev["examples"].append(text[:40])
    summary = {}
    for role, entry in roles.items():
        summary[role] = {
            "paragraphs": entry["paragraphs"],
            "text_formats": [dict(json.loads(k), count=c, example=entry["examples"].get(k, ""))
                             for k, c in entry["runs"].most_common(examples)],
            "paragraph_formats": [dict(json.loads(k), count=c, example=entry["examples"].get(k, ""))
                                  for k, c in entry["layout"].most_common(examples)],
        }
    tables = []
    for table in root.find(W + "body").iter(W + "tbl"):
        if _is_equation_table(table):
            continue
        borders = table.find(f"{W}tblPr/{W}tblBorders")
        desc = {}
        if borders is not None:
            for side in ("top", "bottom", "insideH", "insideV", "left", "right"):
                element = borders.find(W + side)
                if element is not None:
                    value = element.get(W + "val")
                    desc[side] = value if value in ("nil", "none") else f"{value} {int(element.get(W + 'sz', '0')) / 8:g}pt"
        three_line = desc.get("top", "").startswith("single") and desc.get("bottom", "").startswith("single") \
            and desc.get("insideH") in ("nil", "none") and desc.get("insideV") in ("nil", "none")
        tables.append({"rows": len(table.findall(W + "tr")), "borders": desc, "three_line": three_line})
    page = {}
    sect = root.find(f"{W}body/{W}sectPr")
    if sect is not None:
        size, margin = sect.find(W + "pgSz"), sect.find(W + "pgMar")
        if size is not None:
            page["size_cm"] = [round(int(size.get(W + "w")) / 567, 2), round(int(size.get(W + "h")) / 567, 2)]
        if margin is not None:
            page["margins_cm"] = {s: round(int(margin.get(W + s)) / 567, 2) for s in ("top", "bottom", "left", "right")}
    mathtype = sum(1 for _ in root.iter("{urn:schemas-microsoft-com:office:office}OLEObject")
                   if str(_.get("ProgID", "")).startswith("Equation.DSMT"))
    result = {"ok": True, "action": "report-docx-formatting", "document_path": str(path), "page": page,
              "roles": summary, "tables": {"count": len(tables), "three_line": sum(t["three_line"] for t in tables),
                                           "examples": tables[:examples]},
              "mathtype_objects": mathtype,
              "equation_tables": sum(1 for t in root.find(W + "body").iter(W + "tbl") if _is_equation_table(t))}
    if profile:
        page_spec = profile["page"]
        page_issues = []
        if page.get("size_cm") and [round(x, 1) for x in page["size_cm"]] not in (
                [page_spec["width_cm"], page_spec["height_cm"]], [page_spec["height_cm"], page_spec["width_cm"]]):
            page_issues.append(f"page size {page['size_cm']}")
        for side, value in (page.get("margins_cm") or {}).items():
            if abs(value - page_spec[f"{side}_cm"]) > 0.05:
                page_issues.append(f"{side} margin {value} cm")
        result["profile"] = profile.get("name")
        result["deviations"] = {role: {"paragraphs": d["paragraphs"], "issues": dict(d["issues"].most_common(6)),
                                       "examples": d["examples"]} for role, d in deviations.items()}
        if profile.get("tables", {}).get("three_line") and tables and result["tables"]["three_line"] < len(tables):
            page_issues.append(f"{len(tables) - result['tables']['three_line']} table(s) are not three-line tables")
        result["page_and_table_issues"] = page_issues
        result["conforms"] = not deviations and not page_issues
    return result


def _utf8_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass


def main() -> int:
    _utf8_stdio()
    parser = argparse.ArgumentParser()
    sub = parser.add_subparsers(dest="cmd", required=True)
    a = sub.add_parser("apply")
    a.add_argument("input")
    a.add_argument("output")
    a.add_argument("--profile", default="")
    a.add_argument("--overwrite", action="store_true")
    r = sub.add_parser("report")
    r.add_argument("document")
    r.add_argument("--profile", default="")
    args = parser.parse_args()
    try:
        if args.cmd == "apply":
            result = apply(args.input, args.output, args.profile or None, args.overwrite)
        else:
            result = report(args.document, args.profile or None)
    except Exception as exc:
        result = {"ok": False, "action": args.cmd, "error": f"{type(exc).__name__}: {exc}"}
    sys.stdout.write(json.dumps(result, ensure_ascii=False) + "\n")
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
