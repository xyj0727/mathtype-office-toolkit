#!/usr/bin/env python3
"""Find plain-text mathematics in a DOCX and prepare it for MathType rendering.

``scan``    lists every plain-text formula or math symbol (body text and table cells) with its
            context, a suggested MathType TeX string and a proposed action. The result is a JSON
            candidates file that a person or agent can review and edit (change ``action`` or ``tex``).
``prepare`` applies a (reviewed) candidates file: ``mathtype`` candidates become unique
            ``{{MATH:...}}`` markers plus a schema v1 manifest for render_mathtype_word_document;
            ``italic_text`` candidates are rewritten as Times New Roman italic text with real
            Word subscripts (the CJGE rule for simple symbols in running text); ``skip`` is left alone.

Strategies for the proposed action:
  cjge  (default) expressions -> mathtype, single symbols such as m, φ, T_ult -> italic_text
  all             every candidate -> mathtype

Usage: mathtype_scan.py scan DOCX [--candidates OUT.json] [--strategy cjge|all] [--include-references]
       mathtype_scan.py prepare DOCX CANDIDATES.json OUTPUT.docx MANIFEST.json [--overwrite] [--id-prefix eq]
Prints one JSON result line.
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
import os
import re
import sys
import tempfile
import zipfile
from pathlib import Path

WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % WNS
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

GREEK_TEX = {
    "α": r"\alpha", "β": r"\beta", "γ": r"\gamma", "δ": r"\delta", "ε": r"\varepsilon", "ζ": r"\zeta",
    "η": r"\eta", "θ": r"\theta", "ι": r"\iota", "κ": r"\kappa", "λ": r"\lambda", "μ": r"\mu", "ν": r"\nu",
    "ξ": r"\xi", "π": r"\pi", "ρ": r"\rho", "σ": r"\sigma", "τ": r"\tau", "υ": r"\upsilon", "φ": r"\varphi",
    "χ": r"\chi", "ψ": r"\psi", "ω": r"\omega", "Γ": r"\Gamma", "Δ": r"\Delta", "Θ": r"\Theta",
    "Λ": r"\Lambda", "Ξ": r"\Xi", "Π": r"\Pi", "Σ": r"\Sigma", "Φ": r"\Phi", "Ψ": r"\Psi", "Ω": r"\Omega",
}
GREEK = "α-ωΑ-Ω"
# Subscript suffix after a comma, e.g. E50,col  Kp,g  Δσ3,exp  φ_sp
COMMA_SUB = r"(?:,(?:col|sp|exp|max|min|g|p|s|f|u|d|r|t|a|b|e|c))?"
UNITS = r"kN/m|kN·m|kN|kPa|MPa|GPa|Pa|mm|cm|km|m²|m³|m/s|kg/m³|g/cm³|mg|kg|°C|K|s|min|h|m"

TOKEN = re.compile(
    r"(?P<FUNC>(?<![A-Za-z])(?:tan|sin|cos|cot|ln|lg|exp|log)(?=\s*[(A-Za-z" + GREEK + r"]))"
    # letter or Greek with an underscore subscript: T_ult, c_col, K_E0, ε_θ,f, E_p50
    r"|(?P<SYMU>(?<![A-Za-z0-9])(?:Δ?[" + GREEK + r"]|[A-Za-z])_[A-Za-z0-9θ]+(?:,[a-z]+)?)"
    # Greek with a trailing index or short descriptive subscript: σ3, σ1f, Δσ3, ρd,max, ε50%
    r"|(?P<SYMG>Δ?(?!Φ\d)[" + GREEK + r"](?:\d+[a-z]?%?|[a-z]{1,3}(?:,[a-z]{1,3})?(?![a-z]))?" + COMMA_SUB + r")"
    # unit placed before letter rules so g/cm³ is never read as the variable g
    r"|(?P<UNITG>g/cm³|kg/m³)"
    # letter plus digits: E50, R0, d10, K0 (not labels such as T1, T30, P1, P2)
    r"|(?P<SYMD>(?<![A-Za-z0-9])(?!T\d|P\d)[A-Za-z]\d{1,2}[a-z]?" + COMMA_SUB + r"(?![A-Za-z0-9]))"
    # conventional two-letter geotechnical symbols
    r"|(?P<SYMK>(?<![A-Za-z0-9])(?:Dr|Kp|Ka|Kr|Cu|Cc|Ip|Il|Sr)" + COMMA_SUB + r"(?![A-Za-z0-9]))"
    # single letters, excluding a, A, I, i, figure labels (a)-(f) and abbreviations such as e.g.
    r"|(?P<SYM1>(?<![A-Za-z0-9_.–])(?!(?<=\()[a-f]\))[B-HJ-Zb-hj-z](?![A-Za-z0-9_–]|\.[A-Za-z]))"
    r"|(?P<UNIT>(?<![A-Za-z])(?:" + UNITS + r")(?![A-Za-z]))"
    r"|(?P<NUM>\d+(?:\.\d+)?)"
    r"|(?P<OP>[=≈≠∝≤≥<>⇒→+×·/^−√½²³°%–↑±])"
    r"|(?P<LP>[(\[])|(?P<RP>[)\]])|(?P<COMMA>,)|(?P<SP>[ \u00a0]+)"
    r"|(?P<OTHER>.)",
    re.S,
)
SYMBOLS = {"FUNC", "SYMU", "SYMG", "SYMD", "SYMK", "SYM1"}
PREFIX_OPS = ("√", "½", "−", "±")
POSTFIX_OPS = ("²", "³", "°", "%")
OP_TEX = {
    "=": "=", "≈": r"\approx ", "≠": r"\ne ", "∝": r"\propto ", "≤": r"\le ", "≥": r"\ge ", "<": "<", ">": ">",
    "⇒": r"\Rightarrow ", "→": r"\rightarrow ", "+": "+", "×": r"\times ", "·": r"\cdot ", "/": "/", "−": "-",
    "↑": r"\uparrow ", "²": "^{2}", "³": "^{3}", "°": "{}^{\\circ}", "%": r"\%", "½": r"\frac{1}{2}", "±": r"\pm ",
}
DEFAULT_INDEX_SUBSCRIPTS = ("i", "j", "k")


# ---------------------------------------------------------------- tokenising and span detection

def lex(text: str) -> list[list]:
    tokens = [[m.lastgroup, m.group(), m.start(), m.end()] for m in TOKEN.finditer(text)]
    for index, token in enumerate(tokens):
        if token[0] == "UNITG":
            token[0] = "UNIT"
        # the letter m right after a number is the unit metre, not the variable m
        if token[0] == "SYM1" and token[1] in ("m", "s", "h", "K"):
            j = index - 1
            while j >= 0 and tokens[j][0] == "SP":
                j -= 1
            if j >= 0 and tokens[j][0] == "NUM":
                token[0] = "UNIT"
        # a unit word that is not preceded by a number is ordinary text or a variable
        if token[0] == "UNIT" and token[1] not in ("g/cm³", "kg/m³"):
            j = index - 1
            while j >= 0 and tokens[j][0] == "SP":
                j -= 1
            if j < 0 or tokens[j][0] != "NUM":
                token[0] = "SYM1" if re.fullmatch(r"[B-HJ-Zb-hj-z]", token[1]) else "OTHER"
    return tokens


def find_spans(text: str) -> list[tuple[int, int, list]]:
    """Return (start, end, tokens) for every maximal plain-text math span in ``text``."""
    toks = lex(text)
    n = len(toks)

    def prev_idx(i: int) -> int:
        j = i - 1
        while j >= 0 and toks[j][0] == "SP":
            j -= 1
        return j

    def next_idx(i: int) -> int:
        j = i + 1
        while j < n and toks[j][0] == "SP":
            j += 1
        return j

    match: dict[int, int] = {}
    stack: list[int] = []
    for i, token in enumerate(toks):
        if token[0] == "LP":
            stack.append(i)
        elif token[0] == "RP" and stack:
            match[stack.pop()] = i
    rmatch = {b: a for a, b in match.items()}

    def mathy(a: int, b: int) -> bool | None:
        """None: not mathematical; False: only numbers; True: contains a symbol."""
        if a >= b:
            return None
        has_symbol = False
        for j in range(a, b):
            kind = toks[j][0]
            if kind in SYMBOLS:
                has_symbol = True
            elif kind in ("NUM", "OP", "SP", "COMMA", "LP", "RP"):
                pass
            elif kind == "UNIT" and prev_idx(j) >= a and toks[prev_idx(j)][0] == "NUM":
                pass
            else:
                return None
        return has_symbol

    def operand(i: int) -> int | None:
        if i >= n:
            return None
        kind = toks[i][0]
        if kind in SYMBOLS:
            if kind == "FUNC":
                return operand(next_idx(i))
            end = i + 1
            if end < n and toks[end][0] == "LP" and end in match and mathy(end + 1, match[end]) is not None:
                end = match[end] + 1
            return end
        if kind == "NUM":
            end = i + 1
            j = next_idx(i)
            if j < n and toks[j][0] == "UNIT":
                end = j + 1
            return end
        if kind == "LP" and i in match and mathy(i + 1, match[i]) is not None:
            return match[i] + 1
        return None

    def is_seed(i: int) -> bool:
        kind = toks[i][0]
        if kind in SYMBOLS:
            return True
        if kind == "LP" and toks[i][1] == "(" and i in match and bool(mathy(i + 1, match[i])):
            return True
        # pure arithmetic such as 0.16×150 = 24 kPa
        if kind == "NUM" and (i == 0 or toks[i - 1][0] not in ("NUM", "OP")):
            ops = []
            j = i
            while j < n and toks[j][0] in ("NUM", "OP", "SP", "UNIT"):
                if toks[j][0] == "OP":
                    ops.append(toks[j][1])
                j += 1
            return "=" in ops and ("×" in ops or "·" in ops)
        return False

    spans = []
    i = 0
    while i < n:
        if not is_seed(i):
            i += 1
            continue
        start = i
        while True:  # extend left over prefix operators and "number OP" pairs
            j = prev_idx(start)
            if j < 0:
                break
            if toks[j][1] in PREFIX_OPS:
                start = j
                continue
            if toks[j][0] == "OP" and toks[j][1] not in ("–",) + POSTFIX_OPS:
                jj = prev_idx(j)
                if jj >= 0 and toks[jj][0] == "UNIT":
                    jj = prev_idx(jj)
                if jj >= 0 and toks[jj][0] == "NUM":
                    start = jj
                    continue
                if jj >= 0 and toks[jj][0] == "RP" and jj in rmatch and mathy(rmatch[jj] + 1, jj):
                    start = rmatch[jj]
                    continue
            break
        end = None
        p = start
        while p < n:  # extend right over "operand OP operand ..."
            if toks[p][1] in PREFIX_OPS:
                p += 1
                while p < n and toks[p][0] == "SP":
                    p += 1
                continue
            op_end = operand(p)
            if op_end is None:
                break
            end = op_end
            while end < n and toks[end][1] in POSTFIX_OPS:
                end += 1
            if end < n and toks[end][1] == "^":
                p = end + 1
                end = None
                continue
            if end < n and toks[end][0] == "LP" and end in match and mathy(end + 1, match[end]) is not None:
                p = end
                continue
            j = next_idx(end - 1) if end < n else end
            if j < n and toks[j][0] == "OP" and toks[j][1] not in POSTFIX_OPS:
                # "c +25%" in prose: a sign glued to the next number, not an operator
                if toks[j][1] in ("+", "−") and toks[j - 1][0] == "SP" and j + 1 < n and toks[j + 1][0] != "SP":
                    break
                q = j if toks[j][1] in PREFIX_OPS else next_idx(j)
                if q < n and (operand(q) is not None or toks[q][1] in PREFIX_OPS):
                    p = q
                    continue
            break
        if end is None:
            i += 1
            continue
        while toks[end - 1][0] == "SP":
            end -= 1
        spans.append((toks[start][2], toks[end - 1][3], toks[start:end]))
        i = end
    # a placeholder such as [x] is not mathematics
    return [s for s in spans if text[s[0]:s[1]] not in ("x",)]


# ---------------------------------------------------------------- TeX conversion

def _subscript_tex(sub: str, index_subscripts: tuple[str, ...]) -> str:
    out = ""
    for part in re.findall(r"[A-Za-z]+|\d+|θ|,|%", sub):
        if part == "θ":
            out += r"\theta "
        elif part == "%":
            out += r"\%"
        elif part.isdigit() or part == ",":
            out += part
        elif part in index_subscripts:
            out += part
        else:
            out += r"\mathrm{%s}" % part
    return out.strip()


def split_symbol(kind: str, text: str) -> tuple[str, str]:
    """(base, subscript) of a symbol token; base keeps a leading Δ."""
    if kind in ("SYM1", "FUNC"):
        return text, ""
    if kind == "SYMK":
        return text[0], text[1:]
    base = "Δ" + text[1] if text[0] == "Δ" else text[0]
    rest = text[len(base):]
    return base, rest[1:] if rest.startswith("_") else rest


def symbol_tex(kind: str, text: str, index_subscripts: tuple[str, ...]) -> str:
    if kind == "FUNC":
        return "\\" + text + " "
    base, sub = split_symbol(kind, text)
    tex = "".join(GREEK_TEX.get(ch, ch) + (" " if ch in GREEK_TEX else "") for ch in base)
    # keep the space after a trailing control word (eta T), drop it before a subscript
    return tex.rstrip() + "_{%s}" % _subscript_tex(sub, index_subscripts) if sub else tex


def to_tex(tokens: list, index_subscripts: tuple[str, ...] = DEFAULT_INDEX_SUBSCRIPTS) -> str:
    n = len(tokens)

    def one(i: int) -> tuple[str, int]:
        kind, text = tokens[i][0], tokens[i][1]
        if kind in SYMBOLS:
            return symbol_tex(kind, text, index_subscripts), i + 1
        if kind == "NUM":
            return text, i + 1
        if kind == "UNIT":
            unit = text.replace("³", "^{3}").replace("²", "^{2}").replace("°C", "{}^{\\circ}C")
            unit = re.sub(r"([A-Za-z/·]+)", r"\\mathrm{\1}", unit).replace("·", r"\cdot ")
            return r"\ " + unit, i + 1
        if kind == "LP":
            depth = 0
            j = i
            while True:
                if tokens[j][0] == "LP":
                    depth += 1
                if tokens[j][0] == "RP":
                    depth -= 1
                    if depth == 0:
                        break
                j += 1
            inner = to_tex(tokens[i + 1:j], index_subscripts)
            if text == "[":
                return r"\left[" + (inner or r"\ ") + r"\right]", j + 1
            return "(" + inner + ")", j + 1
        if kind == "RP":
            return text, i + 1
        if kind == "COMMA":
            return ",", i + 1
        if kind == "SP":
            return "", i + 1
        if kind == "OP":
            if text in ("√", "^"):
                j = i + 1
                while tokens[j][0] == "SP":
                    j += 1
                argument, after = one(j)
                return (r"\sqrt{%s}" if text == "√" else "^{%s}") % argument.strip(), after
            if text == "–":  # en dash: a range between numbers, otherwise a relation such as E50–σ3
                j = i - 1
                while j >= 0 and tokens[j][0] == "SP":
                    j -= 1
                return (r"\sim " if j >= 0 and tokens[j][0] == "NUM" else "-"), i + 1
            return OP_TEX[text], i + 1
        return text, i + 1

    out = []
    i = 0
    while i < n:
        piece, i = one(i)
        out.append(piece)
    return "".join(out).strip()


def span_kind(tokens: list) -> str:
    """'symbol' for one bare symbol (m, φ, T_ult, E50,col); 'expression' for everything else."""
    meaningful = [t for t in tokens if t[0] != "SP"]
    if len(meaningful) == 1 and meaningful[0][0] in SYMBOLS - {"FUNC"}:
        return "symbol"
    return "expression"


# ---------------------------------------------------------------- DOCX traversal

def _heading_levels(styles_xml: bytes | None) -> dict[str, int]:
    from lxml import etree

    levels: dict[str, int] = {}
    if not styles_xml:
        return levels
    root = etree.fromstring(styles_xml)
    for style in root.findall(f"{W}style"):
        style_id = style.get(f"{W}styleId")
        name = style.find(f"{W}name")
        name_value = (name.get(f"{W}val") if name is not None else "") or ""
        match = re.fullmatch(r"(?i)heading\s*(\d)", name_value.strip())
        outline = style.find(f"{W}pPr/{W}outlineLvl")
        if match:
            levels[style_id] = int(match.group(1))
        elif outline is not None:
            levels[style_id] = int(outline.get(f"{W}val")) + 1
    return levels


REFERENCE_HEADING = re.compile(r"^\s*[\d.]*\s*(参考文献|參考文獻|References|Bibliography|Literature cited)\b", re.I)


def _paragraph_text_nodes(paragraph) -> list:
    """w:t nodes of a paragraph in reading order (runs directly in the paragraph, hyperlinks, smart tags)."""
    nodes = []
    for run in paragraph.iter(f"{W}r"):
        # skip runs that belong to a nested paragraph (text boxes) — they are visited on their own
        if run.getparent() is not paragraph and _owning_paragraph(run) is not paragraph:
            continue
        for child in run:
            if child.tag == f"{W}t":
                nodes.append(child)
    return nodes


def _owning_paragraph(element):
    parent = element.getparent()
    while parent is not None and parent.tag != f"{W}p":
        parent = parent.getparent()
    return parent


def iter_paragraphs(root, levels: dict[str, int], include_references: bool):
    """Yield (index, paragraph, in_table) for body paragraphs, skipping the reference list by default."""
    skipping_level = None
    for index, paragraph in enumerate(root.iter(f"{W}p")):
        style = paragraph.find(f"{W}pPr/{W}pStyle")
        level = levels.get(style.get(f"{W}val")) if style is not None else None
        text = "".join(t.text or "" for t in _paragraph_text_nodes(paragraph))
        if skipping_level is not None:
            if level is not None and level <= skipping_level:
                skipping_level = None
            else:
                continue
        if not include_references and level is not None and REFERENCE_HEADING.match(text):
            skipping_level = level
            continue
        in_table = any(ancestor.tag == f"{W}tc" for ancestor in paragraph.iterancestors())
        yield index, paragraph, in_table


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def scan(document_path: str, strategy: str = "cjge", include_references: bool = False,
         index_subscripts: tuple[str, ...] = DEFAULT_INDEX_SUBSCRIPTS) -> dict:
    from lxml import etree

    if strategy not in ("cjge", "all"):
        raise ValueError("strategy must be 'cjge' or 'all'")
    path = Path(document_path).resolve()
    with zipfile.ZipFile(path) as package:
        root = etree.fromstring(package.read("word/document.xml"))
        levels = _heading_levels(package.read("word/styles.xml") if "word/styles.xml" in package.namelist() else None)
    candidates = []
    for index, paragraph, in_table in iter_paragraphs(root, levels, include_references):
        text = "".join(t.text or "" for t in _paragraph_text_nodes(paragraph))
        if not text.strip() or "{{MATH:" in text or "{{EQREF:" in text:
            continue
        for start, end, tokens in find_spans(text):
            kind = span_kind(tokens)
            action = "mathtype" if strategy == "all" or kind == "expression" else "italic_text"
            candidates.append({
                "id": f"c{len(candidates) + 1:04d}",
                "paragraph_index": index,
                "start": start,
                "end": end,
                "text": text[start:end],
                "context": text[max(0, start - 20):end + 20],
                "in_table": in_table,
                "kind": kind,
                "tex": to_tex(tokens, index_subscripts),
                "action": action,
            })
    counts = {a: sum(1 for c in candidates if c["action"] == a) for a in ("mathtype", "italic_text", "skip")}
    return {
        "ok": True,
        "action": "scan",
        "schema": "mathtype-for-word-candidates/1",
        "document_path": str(path),
        "document_sha256": _sha256(path),
        "strategy": strategy,
        "include_references": include_references,
        "candidates_found": len(candidates),
        "counts": counts,
        "candidates": candidates,
    }


# ---------------------------------------------------------------- applying candidates

def _replace_text_range(nodes: list, start: int, end: int, replacement: str) -> None:
    """Replace characters [start, end) of the concatenated w:t nodes; the replacement goes into the first node."""
    offset = 0
    placed = False
    for node in nodes:
        text = node.text or ""
        lo, hi = max(start, offset), min(end, offset + len(text))
        if lo < hi or (not placed and lo == hi == start and offset <= start <= offset + len(text) and start == end):
            node.text = text[:lo - offset] + ("" if placed else replacement) + text[hi - offset:]
            node.set(XML_SPACE, "preserve")
            placed = True
        offset += len(text)


def _italic_segments(tokens: list, index_subscripts: tuple[str, ...]) -> list[tuple[str, bool, str | None]]:
    """(text, italic, vertAlign) runs for a bare symbol, CJGE style: italic base, upright descriptive subscript."""
    token = next(t for t in tokens if t[0] != "SP")
    base, sub = split_symbol(token[0], token[1])
    segments: list[tuple[str, bool, str | None]] = []
    if base.startswith("Δ"):
        segments.append(("Δ", False, None))
        base = base[1:]
    segments.append((base, True, None))
    if sub:
        for part in re.findall(r"[A-Za-z]+|[^A-Za-z]+", sub):
            segments.append((part, part in index_subscripts, "subscript"))
    return segments


def _apply_italic(paragraph, start: int, end: int, segments) -> bool:
    """Rewrite [start, end) as formatted runs. Only spans inside one single-w:t run are rewritten."""
    from lxml import etree

    offset = 0
    for node in _paragraph_text_nodes(paragraph):
        text = node.text or ""
        if offset <= start and end <= offset + len(text):
            run = node.getparent()
            if sum(1 for child in run if child.tag == f"{W}t") != 1:
                return False
            before, after = text[:start - offset], text[end - offset:]
            rpr = run.find(f"{W}rPr")
            parent = run.getparent()
            position = parent.index(run)
            new_runs = []

            def make(content: str, italic: bool | None, vert: str | None):
                new = etree.Element(f"{W}r")
                props = copy.deepcopy(rpr) if rpr is not None else etree.Element(f"{W}rPr")
                if italic is not None:
                    fonts = props.find(f"{W}rFonts")
                    if fonts is None:
                        fonts = etree.Element(f"{W}rFonts")
                        props.insert(0, fonts)
                    for attr in ("ascii", "hAnsi", "cs"):
                        fonts.set(f"{W}{attr}", "Times New Roman")
                    for tag in ("i", "iCs"):
                        old = props.find(f"{W}{tag}")
                        if old is not None:
                            props.remove(old)
                        if italic:
                            etree.SubElement(props, f"{W}{tag}")
                    old_vert = props.find(f"{W}vertAlign")
                    if old_vert is not None:
                        props.remove(old_vert)
                    if vert:
                        etree.SubElement(props, f"{W}vertAlign").set(f"{W}val", vert)
                if len(props):
                    new.append(props)
                t = etree.SubElement(new, f"{W}t")
                t.text = content
                t.set(XML_SPACE, "preserve")
                return new

            if before:
                new_runs.append(make(before, None, None))
            for content, italic, vert in segments:
                new_runs.append(make(content, italic, vert))
            if after:
                new_runs.append(make(after, None, None))
            parent.remove(run)
            for k, new in enumerate(new_runs):
                parent.insert(position + k, new)
            return True
        offset += len(text)
    return False


def prepare(document_path: str, candidates_path: str, output_path: str, manifest_path: str,
            overwrite: bool = False, id_prefix: str = "eq",
            index_subscripts: tuple[str, ...] = DEFAULT_INDEX_SUBSCRIPTS) -> dict:
    from lxml import etree

    source = Path(document_path).resolve()
    output = Path(output_path).resolve()
    manifest_file = Path(manifest_path).resolve()
    for target in (output, manifest_file):
        if target.exists() and not overwrite:
            raise FileExistsError(f"Output exists (pass overwrite): {target}")
    if output == source:
        raise ValueError("Write the prepared DOCX to a new path; the source is preserved.")
    data = json.loads(Path(candidates_path).read_text(encoding="utf-8-sig"))
    if data.get("document_sha256") and data["document_sha256"] != _sha256(source):
        raise ValueError("The DOCX changed after it was scanned; scan it again.")
    if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]*", id_prefix):
        raise ValueError("id_prefix must be an ASCII identifier.")
    with zipfile.ZipFile(source) as package:
        root = etree.fromstring(package.read("word/document.xml"))
        paragraphs = list(root.iter(f"{W}p"))
        by_paragraph: dict[int, list[dict]] = {}
        for candidate in data.get("candidates", []):
            if candidate.get("action") in ("mathtype", "italic_text"):
                by_paragraph.setdefault(int(candidate["paragraph_index"]), []).append(candidate)
        ordered = sorted((c for group in by_paragraph.values() for c in group),
                         key=lambda c: (int(c["paragraph_index"]), int(c["start"])))
        number = {c["id"]: k + 1 for k, c in enumerate(c for c in ordered if c["action"] == "mathtype")}
        width = max(3, len(str(len(number))))
        equations, italic_done, problems = [], 0, []
        for index, group in by_paragraph.items():
            paragraph = paragraphs[index]
            text = "".join(t.text or "" for t in _paragraph_text_nodes(paragraph))
            for candidate in sorted(group, key=lambda c: int(c["start"]), reverse=True):
                start, end = int(candidate["start"]), int(candidate["end"])
                if text[start:end] != candidate["text"]:
                    problems.append(f"{candidate['id']}: text no longer matches ({candidate['text']!r}).")
                    continue
                if candidate["action"] == "mathtype":
                    eq_id = f"{id_prefix}{number[candidate['id']]:0{width}d}"
                    marker = "{{MATH:%s}}" % eq_id
                    _replace_text_range(_paragraph_text_nodes(paragraph), start, end, marker)
                    equations.append({"id": eq_id, "marker": marker, "tex": candidate["tex"],
                                      "layout": "inline", "numbered": False, "source_text": candidate["text"]})
                else:
                    tokens = lex(candidate["text"])
                    if span_kind(tokens) != "symbol" or not _apply_italic(
                            paragraph, start, end, _italic_segments(tokens, index_subscripts)):
                        problems.append(f"{candidate['id']}: italic_text needs one bare symbol inside one run; left as is.")
                        continue
                    italic_done += 1
                text = "".join(t.text or "" for t in _paragraph_text_nodes(paragraph))
        equations.sort(key=lambda e: e["id"])
        body = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        output.parent.mkdir(parents=True, exist_ok=True)
        handle, temporary = tempfile.mkstemp(suffix=".docx", dir=output.parent)
        os.close(handle)
        try:
            with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as out:
                for item in package.infolist():
                    out.writestr(item, body if item.filename == "word/document.xml" else package.read(item.filename))
        except BaseException:
            os.unlink(temporary)
            raise
    os.replace(temporary, output)
    manifest = {"schema_version": 1, "equations": equations, "references": []}
    manifest_file.write_text(json.dumps(manifest, ensure_ascii=False, indent=1), encoding="utf-8")
    return {"ok": not problems, "action": "prepare", "input_path": str(source), "output_path": str(output),
            "manifest_path": str(manifest_file), "mathtype_markers": len(equations),
            "italic_text_symbols": italic_done, "problems": problems}


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
    s = sub.add_parser("scan")
    s.add_argument("document")
    s.add_argument("--candidates", default="")
    s.add_argument("--strategy", default="cjge", choices=("cjge", "all"))
    s.add_argument("--include-references", action="store_true")
    s.add_argument("--index-subscripts", default="i,j,k")
    p = sub.add_parser("prepare")
    p.add_argument("document")
    p.add_argument("candidates")
    p.add_argument("output")
    p.add_argument("manifest")
    p.add_argument("--overwrite", action="store_true")
    p.add_argument("--id-prefix", default="eq")
    p.add_argument("--index-subscripts", default="i,j,k")
    args = parser.parse_args()
    index_subscripts = tuple(x for x in args.index_subscripts.split(",") if x)
    try:
        if args.cmd == "scan":
            result = scan(args.document, args.strategy, args.include_references, index_subscripts)
            if args.candidates:
                Path(args.candidates).write_text(json.dumps(result, ensure_ascii=False, indent=1), encoding="utf-8")
                result = {k: v for k, v in result.items() if k != "candidates"}
                result["candidates_path"] = str(Path(args.candidates).resolve())
                result["examples"] = []
        else:
            result = prepare(args.document, args.candidates, args.output, args.manifest,
                             args.overwrite, args.id_prefix, index_subscripts)
    except Exception as exc:
        result = {"ok": False, "action": args.cmd, "error": f"{type(exc).__name__}: {exc}"}
    sys.stdout.write(json.dumps(result, ensure_ascii=False) + "\n")
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
