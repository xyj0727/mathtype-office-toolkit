#!/usr/bin/env python3
"""Apply the table display layout to a rendered DOCX (default profile: CJGE, config/cjge_layout_profile.json;
the word-mathtype-mcp profile config/repo_format_profile.json can be passed with --profile).

Layout (mirrors paper_core/export/word_postprocessor._apply_formula_paragraph):
  * each numbered/unnumbered MathType display paragraph becomes a 1x3 borderless table;
  * side cells formula_side_cell_width_pt (72 pt, capped at text width / 4), middle cell = remainder;
  * zero cell padding, row height "at least" formula line spacing, cells vertically centred;
  * cell paragraphs: no indents, 0 pt before/after, "at least" formula_placeholder_line_spacing_pt;
  * equation centred in the middle cell, number right-aligned in the right cell in the body font.

Unlike the upstream implementation this converts the existing tab-separated MathType paragraph in place
(Range.ConvertToTable), so the Equation.DSMT4 object, the MathType MTPlaceRef/SEQ number field and its
ZEqnNum bookmark stay intact and MathType references keep resolving. No clipboard or UI automation.

Usage: repo_layout.py INPUT OUTPUT [--profile PATH] [--overwrite]  -> prints one JSON result line.
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
from pathlib import Path

DEFAULT_PROFILE = Path(__file__).resolve().parent.parent / "config" / "cjge_layout_profile.json"

WD_SEPARATE_BY_TABS = 1
WD_ALIGN_CENTER = 1
WD_ALIGN_RIGHT = 2
WD_LINE_SPACE_AT_LEAST = 3
WD_ROW_HEIGHT_AT_LEAST = 1
WD_CELL_ALIGN_VERTICAL_CENTER = 1
WD_LINE_STYLE_NONE = 0
WD_WITHIN_TABLE = 12
WD_FORMAT_XML_DOCUMENT = 16


def _float(profile: dict, key: str, default: float) -> float:
    try:
        return float(profile.get(key, default))
    except (TypeError, ValueError):
        return float(default)


def _is_mathtype_display(paragraph) -> bool:
    rng = paragraph.Range
    if rng.Information(WD_WITHIN_TABLE):
        return False
    if rng.InlineShapes.Count != 1:
        return False
    shape = rng.InlineShapes(1)
    try:
        prog_id = str(shape.OLEFormat.ProgID)
    except Exception:
        return False
    if not prog_id.startswith("Equation.DSMT"):
        return False
    style = str(paragraph.Style.NameLocal) if paragraph.Style is not None else ""
    text = rng.Text
    # MathType display paragraphs are "<tab><object><tab><number field>" (number optional).
    return style == "MTDisplayEquation" or text.startswith("\t")


def apply_layout(input_path: str, output_path: str, profile_path: str = "", overwrite: bool = False) -> dict:
    import pythoncom
    import win32com.client

    src = Path(input_path).resolve()
    dst = Path(output_path).resolve()
    if not src.exists():
        raise FileNotFoundError(f"input not found: {src}")
    if src != dst and dst.exists() and not overwrite:
        raise FileExistsError(f"output exists: {dst} (pass overwrite=true)")
    if src == dst and not overwrite:
        raise ValueError("output equals input; pass overwrite=true to edit in place")
    profile = json.loads(Path(profile_path or DEFAULT_PROFILE).read_text(encoding="utf-8-sig"))

    work = dst.with_name(f".{dst.stem}.repo-layout.tmp{dst.suffix}")
    shutil.copyfile(src, work)

    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word_pid = None
    try:
        import win32process

        word_pid = win32process.GetWindowThreadProcessId(word.Hwnd)[1] if hasattr(word, "Hwnd") else None
    except Exception:
        pass
    converted, skipped = 0, []
    try:
        word.Visible = False
        word.DisplayAlerts = 0
        word.ScreenUpdating = False
        doc = word.Documents.Open(str(work), False, False, False)
        setup = doc.PageSetup
        usable = float(setup.PageWidth - setup.LeftMargin - setup.RightMargin)
        side = min(_float(profile, "formula_side_cell_width_pt", 72.0), max(36.0, usable / 4))
        middle = max(usable - side * 2, usable / 2)
        line_pt = _float(profile, "formula_placeholder_line_spacing_pt", 20.0)
        body_font = str(profile.get("body_font_name", "")).strip()
        east_font = str(profile.get("body_east_asia_font_name", "")).strip()
        body_size = _float(profile, "body_font_size_pt", 12.0)

        for index in range(doc.Paragraphs.Count, 0, -1):
            paragraph = doc.Paragraphs(index)
            if not _is_mathtype_display(paragraph):
                continue
            rng = paragraph.Range
            body = doc.Range(rng.Start, rng.End - 1)  # exclude paragraph mark
            tab_count = body.Text.count("\t")
            if tab_count == 1:
                body.InsertAfter("\t")  # unnumbered display: add an empty number column
            elif tab_count != 2:
                skipped.append({"paragraph": index, "reason": f"expected 1-2 tabs, found {tab_count}"})
                continue
            if body.Text.startswith("\t") is False:
                skipped.append({"paragraph": index, "reason": "does not start with a tab"})
                continue
            table = doc.Range(rng.Start, rng.End).ConvertToTable(WD_SEPARATE_BY_TABS, 1, 3)
            table.AllowAutoFit = False
            table.LeftPadding = 0
            table.RightPadding = 0
            table.TopPadding = 0
            table.BottomPadding = 0
            table.Borders.Enable = False
            for border_index in range(1, 7):
                table.Borders(-border_index).LineStyle = WD_LINE_STYLE_NONE
            row = table.Rows(1)
            row.AllowBreakAcrossPages = False
            row.HeightRule = WD_ROW_HEIGHT_AT_LEAST
            row.Height = line_pt
            table.Cell(1, 1).Width = side
            table.Cell(1, 2).Width = middle
            table.Cell(1, 3).Width = side
            for cell_index in (1, 2, 3):
                cell = table.Cell(1, cell_index)
                cell.VerticalAlignment = WD_CELL_ALIGN_VERTICAL_CENTER
                fmt = cell.Range.ParagraphFormat
                fmt.TabStops.ClearAll()
                fmt.LeftIndent = 0
                fmt.RightIndent = 0
                fmt.FirstLineIndent = 0
                fmt.CharacterUnitFirstLineIndent = 0
                fmt.SpaceBefore = 0
                fmt.SpaceAfter = 0
                fmt.LineSpacingRule = WD_LINE_SPACE_AT_LEAST
                fmt.LineSpacing = line_pt
            table.Cell(1, 2).Range.ParagraphFormat.Alignment = WD_ALIGN_CENTER
            number_range = table.Cell(1, 3).Range
            number_range.ParagraphFormat.Alignment = WD_ALIGN_RIGHT
            if body_font:
                number_range.Font.Name = body_font
            if east_font:
                number_range.Font.NameFarEast = east_font
            if body_size > 0:
                number_range.Font.Size = body_size
            converted += 1

        doc.Fields.Update()
        doc.SaveAs2(str(work), WD_FORMAT_XML_DOCUMENT)
        doc.Close(0)
    finally:
        try:
            word.Quit(0)
        except Exception:
            pass
        pythoncom.CoUninitialize()
    os.replace(work, dst)
    return {
        "ok": True,
        "action": "apply-repo-layout",
        "input_path": str(src),
        "output_path": str(dst),
        "profile": str(Path(profile_path or DEFAULT_PROFILE).resolve()),
        "converted_display_equations": converted,
        "skipped": skipped,
        "layout": {
            "type": "borderless 1x3 table",
            "side_cell_width_pt": round(side, 2),
            "middle_cell_width_pt": round(middle, 2),
            "line_spacing_at_least_pt": line_pt,
            "number_font": {"latin": body_font, "east_asia": east_font, "size_pt": body_size},
        },
        "word_pid": word_pid,
    }



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
    parser.add_argument("input")
    parser.add_argument("output")
    parser.add_argument("--profile", default="")
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()
    try:
        result = apply_layout(args.input, args.output, args.profile, args.overwrite)
    except Exception as exc:  # report as JSON for the MCP wrapper
        result = {"ok": False, "action": "apply-repo-layout", "error": f"{type(exc).__name__}: {exc}"}
    sys.stdout.write(json.dumps(result, ensure_ascii=False) + "\n")
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
