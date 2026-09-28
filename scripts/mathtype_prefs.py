#!/usr/bin/env python3
"""Re-typeset every MathType equation in a DOCX with a MathType preference file (.eqp).

This is what MathType's own "Format Equations" command does, performed silently at file level:

* each ``Equation.DSMT4`` object's MTEF (the ``Equation Native`` stream) is transformed by
  MathType's API (``MTXFormSetPrefs`` + ``MTXFormEqn`` in MathPage.wll) with the preferences;
* MathType renders a new WMF preview, and the object's display size and baseline offset are
  updated from that preview;
* the preferences are stored in the document (``MTPreferences`` custom properties), so equations
  that the user inserts later in MathType use the same settings.

No clipboard, no dialogs, no visible windows: the DOCX package is edited directly, so numbers,
references, fields and bookmarks are untouched.

Usage: mathtype_prefs.py docx INPUT OUTPUT [--preferences EQP] [--overwrite]
       mathtype_prefs.py ole-dir DIRECTORY [--preferences EQP]   (internal: PPTX conversion helper)
Prints one JSON result line.
"""
from __future__ import annotations

import argparse
import ctypes
import json
import os
import re
import shutil
import struct
import sys
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_PREFERENCES = ROOT / "config" / "cjge_equation_preferences.eqp"

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "cp": "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
    "vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
}
CUSTOM_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
CUSTOM_CT = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"

# MathType API constants (MathType Commands template, MTAPI).
MTXFM_PREF_USER = 3
MTXFM_LOCAL = -3
MTXFM_FILE = -4
MTXFM_MTEF = 4
MTXFM_PICT = 6
MTXFM_STAT_ACTUAL_LEN = -1
STGM_READWRITE_EXCLUSIVE = 0x12


def mathpage_path() -> Path:
    import winreg

    candidates = []
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"CLSID\{0002CE03-0000-0000-C000-000000000046}\LocalServer32") as key:
            exe = str(winreg.QueryValueEx(key, "")[0]).strip().strip('"').split('"')[0]
            candidates.append(Path(exe).parent / "MathPage" / ("64" if sys.maxsize > 2**32 else "32") / "MathPage.wll")
    except OSError:
        pass
    candidates.append(Path(r"C:\Program Files (x86)\MathType\MathPage") / ("64" if sys.maxsize > 2**32 else "32") / "MathPage.wll")
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    raise FileNotFoundError("MathPage.wll (MathType 7 API) was not found; install desktop MathType for Windows.")


class MathTypeApi:
    """Thin ctypes wrapper over the MathType API exported by MathPage.wll."""

    def __init__(self) -> None:
        self.dll = ctypes.WinDLL(str(mathpage_path()))
        self.dll.MTInitAPI.argtypes = [ctypes.c_short, ctypes.c_short]
        self.dll.MTInitAPI.restype = ctypes.c_long
        self.dll.MTXFormSetPrefs.argtypes = [ctypes.c_short, ctypes.c_char_p]
        self.dll.MTXFormSetPrefs.restype = ctypes.c_long
        self.dll.MTXFormGetStatus.argtypes = [ctypes.c_short]
        self.dll.MTXFormGetStatus.restype = ctypes.c_long
        self.dll.MTGetPrefsFromFile.argtypes = [ctypes.c_char_p, ctypes.c_char_p, ctypes.c_short]
        self.dll.MTGetPrefsFromFile.restype = ctypes.c_long
        self.dll.MTXFormEqn.argtypes = [
            ctypes.c_short, ctypes.c_short, ctypes.c_char_p, ctypes.c_long,
            ctypes.c_short, ctypes.c_short, ctypes.c_char_p, ctypes.c_long,
            ctypes.c_char_p, ctypes.c_void_p,
        ]
        self.dll.MTXFormEqn.restype = ctypes.c_long

    def __enter__(self) -> "MathTypeApi":
        rc = self.dll.MTInitAPI(0, 30)
        if rc != 1:
            raise RuntimeError(f"MathType API initialisation failed (MTInitAPI returned {rc}).")
        return self

    def __exit__(self, *exc) -> None:
        try:
            self.dll.MTTermAPI()
        except Exception:
            pass

    def preferences_from_file(self, eqp: Path) -> bytes:
        path = str(eqp).encode("mbcs")
        size = self.dll.MTGetPrefsFromFile(path, None, 0)
        if size <= 0:
            raise RuntimeError(f"MathType could not read preferences from {eqp} (code {size}).")
        buffer = ctypes.create_string_buffer(size + 16)
        rc = self.dll.MTGetPrefsFromFile(path, buffer, len(buffer))
        if rc != 0:
            raise RuntimeError(f"MathType could not read preferences from {eqp} (code {rc}).")
        return buffer.value

    def _set_prefs(self, prefs: bytes) -> None:
        self.dll.MTXFormReset()
        rc = self.dll.MTXFormSetPrefs(MTXFM_PREF_USER, prefs)
        if rc != 0:
            raise RuntimeError(f"MTXFormSetPrefs failed with code {rc}.")

    def restyle(self, mtef: bytes, prefs: bytes) -> bytes:
        self._set_prefs(prefs)
        out = ctypes.create_string_buffer(max(4096, len(mtef) * 4))
        rc = self.dll.MTXFormEqn(MTXFM_LOCAL, MTXFM_MTEF, mtef, len(mtef), MTXFM_LOCAL, MTXFM_MTEF,
                                 out, len(out), b" ", None)
        if rc != 0:
            raise RuntimeError(f"MTXFormEqn (MTEF) failed with code {rc}.")
        length = self.dll.MTXFormGetStatus(MTXFM_STAT_ACTUAL_LEN)
        if length <= 0 or length > len(out):
            raise RuntimeError(f"MTXFormEqn returned an invalid MTEF length {length}.")
        return out.raw[:length]

    def render_wmf(self, mtef: bytes, prefs: bytes, target: Path) -> None:
        self._set_prefs(prefs)
        if target.exists():
            target.unlink()
        rc = self.dll.MTXFormEqn(MTXFM_LOCAL, MTXFM_MTEF, mtef, len(mtef), MTXFM_FILE, MTXFM_PICT,
                                 None, 0, str(target).encode("mbcs"), None)
        if rc != 0 or not target.is_file():
            raise RuntimeError(f"MTXFormEqn (WMF) failed with code {rc}.")


def wmf_size_points(data: bytes) -> tuple[float, float]:
    key, _hmf, left, top, right, bottom, inch = struct.unpack_from("<IHhhhhH", data, 0)
    if key != 0x9AC6CDD7 or inch == 0:
        raise ValueError("MathType preview is not a placeable WMF.")
    return (right - left) / inch * 72.0, (bottom - top) / inch * 72.0


def read_equation_native(ole_path: Path) -> tuple[bytes, bytes]:
    import pythoncom

    storage = pythoncom.StgOpenStorage(str(ole_path), None, STGM_READWRITE_EXCLUSIVE, None, 0)
    stream = storage.OpenStream("Equation Native", None, STGM_READWRITE_EXCLUSIVE, 0)
    data = b""
    while True:
        chunk = stream.Read(65536)
        if not chunk:
            break
        data += chunk
    del stream, storage
    header_len, _version, _cf, object_len = struct.unpack_from("<HIHI", data, 0)
    return data[:header_len], data[header_len:header_len + object_len]


def write_equation_native(ole_path: Path, header: bytes, mtef: bytes) -> None:
    import pythoncom

    header = bytearray(header)
    struct.pack_into("<I", header, 8, len(mtef))  # EQNOLEFILEHDR.cbObject
    payload = bytes(header) + mtef
    storage = pythoncom.StgOpenStorage(str(ole_path), None, STGM_READWRITE_EXCLUSIVE, None, 0)
    stream = storage.OpenStream("Equation Native", None, STGM_READWRITE_EXCLUSIVE, 0)
    stream.SetSize(len(payload))
    stream.Seek(0, 0)
    stream.Write(payload)
    stream.Commit(0)
    del stream
    storage.Commit(0)
    del storage


def restyle_ole_file(api: MathTypeApi, prefs: bytes, ole_path: Path, wmf_path: Path) -> tuple[float, float]:
    header, mtef = read_equation_native(ole_path)
    new_mtef = api.restyle(mtef, prefs)
    write_equation_native(ole_path, header, new_mtef)
    api.render_wmf(new_mtef, prefs, wmf_path)
    return wmf_size_points(wmf_path.read_bytes())


# --------------------------------------------------------------------------- DOCX


def _rels_map(zf: zipfile.ZipFile, part: str) -> tuple[str, dict]:
    folder, name = part.rsplit("/", 1)
    rels_name = f"{folder}/_rels/{name}.rels"
    if rels_name not in zf.namelist():
        return rels_name, {}
    from lxml import etree

    root = etree.fromstring(zf.read(rels_name))
    return rels_name, {rel.get("Id"): rel.get("Target") for rel in root.findall("rel:Relationship", NS)}


def _resolve(part: str, target: str) -> str:
    folder = part.rsplit("/", 1)[0]
    parts = f"{folder}/{target}".split("/")
    out = []
    for piece in parts:
        if piece == "..":
            out.pop()
        elif piece and piece != ".":
            out.append(piece)
    return "/".join(out)


def _fmt(value: float) -> str:
    return f"{value:.2f}".rstrip("0").rstrip(".")


# VML shape styles may use any CSS length unit; Word writes exactly 72 pt as "1in".
_CSS_UNIT_PT = {"pt": 1.0, "in": 72.0, "cm": 72.0 / 2.54, "mm": 72.0 / 25.4, "pc": 12.0, "px": 0.75}
_CSS_LENGTH = r"(?<![-\w]){prop}:\s*(-?[\d.]+)\s*(pt|in|cm|mm|pc|px)?"


def _css_length_pt(style: str, prop: str) -> float | None:
    match = re.search(_CSS_LENGTH.format(prop=prop), style)
    if not match:
        return None
    return float(match.group(1)) * _CSS_UNIT_PT[match.group(2) or "px"]


def _set_css_length_pt(style: str, prop: str, value: float) -> str:
    return re.sub(_CSS_LENGTH.format(prop=prop), f"{prop}:{_fmt(value)}pt", style)


def _custom_properties_xml(existing: bytes | None, prefs_text: str) -> bytes:
    from lxml import etree

    if existing:
        root = etree.fromstring(existing)
    else:
        root = etree.Element(f"{{{NS['cp']}}}Properties", nsmap={None: NS["cp"], "vt": NS["vt"]})
    for prop in list(root):
        name = prop.get("name") or ""
        if name == "MTUseMTPrefs" or re.fullmatch(r"MTPreferences( \d+)?", name) or name == "MTPreferenceSource":
            root.remove(prop)
    pids = [int(p.get("pid")) for p in root if (p.get("pid") or "").isdigit()]
    pid = max(pids + [1]) + 1
    chunks = [prefs_text[i:i + 255] for i in range(0, len(prefs_text), 255)]
    items = [("MTPreferences" if i == 0 else f"MTPreferences {i}", chunk) for i, chunk in enumerate(chunks)]
    items.append(("MTPreferenceSource", "cjge_equation_preferences.eqp"))
    for name, value in items:
        prop = etree.SubElement(root, f"{{{NS['cp']}}}property", fmtid=FMTID, pid=str(pid), name=name)
        etree.SubElement(prop, f"{{{NS['vt']}}}lpwstr").text = value
        pid += 1
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def apply_to_docx(input_path: str, output_path: str, preferences: str = "", overwrite: bool = False,
                  store_document_preferences: bool = True) -> dict:
    from lxml import etree

    src, dst = Path(input_path).resolve(), Path(output_path).resolve()
    if not src.is_file():
        raise FileNotFoundError(f"input not found: {src}")
    if src != dst and dst.exists() and not overwrite:
        raise FileExistsError(f"output exists: {dst} (pass overwrite=true)")
    if src == dst and not overwrite:
        raise ValueError("output equals input; pass overwrite=true to edit in place")
    eqp = Path(preferences).resolve() if preferences else DEFAULT_PREFERENCES
    work = Path(tempfile.mkdtemp(prefix="mathtype-prefs-"))
    try:
        zin = zipfile.ZipFile(src)
        replaced: dict[str, bytes] = {}
        added_content_types: dict[str, str] = {}
        report = []
        with MathTypeApi() as api:
            prefs = api.preferences_from_file(eqp)
            xml_parts = [n for n in zin.namelist()
                         if re.fullmatch(r"word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml", n)]
            for part in xml_parts:
                root = etree.fromstring(zin.read(part))
                objects = root.findall(".//w:object", NS)
                if not objects:
                    continue
                _rels_name, rels = _rels_map(zin, part)
                changed = False
                for obj in objects:
                    ole = obj.find("o:OLEObject", NS)
                    shape = obj.find("v:shape", NS)
                    if ole is None or shape is None or not str(ole.get("ProgID", "")).startswith("Equation.DSMT"):
                        continue
                    image = shape.find("v:imagedata", NS)
                    ole_part = _resolve(part, rels[ole.get(f"{{{NS['r']}}}id")])
                    image_part = _resolve(part, rels[image.get(f"{{{NS['r']}}}id")]) if image is not None else None
                    ole_file = work / Path(ole_part).name
                    ole_file.write_bytes(replaced.get(ole_part) or zin.read(ole_part))
                    wmf_file = work / (Path(ole_part).stem + ".wmf")
                    style = shape.get("style", "")
                    old_h = _css_length_pt(style, "height")
                    width, height = restyle_ole_file(api, prefs, ole_file, wmf_file)
                    replaced[ole_part] = ole_file.read_bytes()
                    if image_part:
                        if not image_part.lower().endswith(".wmf"):
                            raise RuntimeError(f"Unsupported equation preview format: {image_part}")
                        replaced[image_part] = wmf_file.read_bytes()
                    style = _set_css_length_pt(style, "width", width)
                    style = _set_css_length_pt(style, "height", height)
                    shape.set("style", style)
                    obj.set(f"{{{NS['w']}}}dxaOrig", str(round(width * 20)))
                    obj.set(f"{{{NS['w']}}}dyaOrig", str(round(height * 20)))
                    run = obj.getparent()
                    position = run.find("w:rPr/w:position", NS) if run is not None else None
                    if position is not None and old_h:
                        old = int(position.get(f"{{{NS['w']}}}val"))
                        position.set(f"{{{NS['w']}}}val", str(round(old * height / old_h)))
                    report.append({"part": part, "object": ole_part, "size_pt": [round(width, 2), round(height, 2)]})
                    changed = True
                if changed:
                    replaced[part] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

            if store_document_preferences:
                prefs_text = prefs.decode("mbcs")
                custom = "docProps/custom.xml"
                existing = zin.read(custom) if custom in zin.namelist() else None
                replaced[custom] = _custom_properties_xml(existing, prefs_text)
                if existing is None:
                    rels_root = etree.fromstring(zin.read("_rels/.rels"))
                    ids = {r.get("Id") for r in rels_root}
                    new_id = next(f"rId{i}" for i in range(1, 1000) if f"rId{i}" not in ids)
                    etree.SubElement(rels_root, f"{{{NS['rel']}}}Relationship", Id=new_id, Type=CUSTOM_REL,
                                     Target="docProps/custom.xml")
                    replaced["_rels/.rels"] = etree.tostring(rels_root, xml_declaration=True, encoding="UTF-8",
                                                             standalone=True)
                    added_content_types["/docProps/custom.xml"] = CUSTOM_CT

        if added_content_types:
            ct_root = etree.fromstring(zin.read("[Content_Types].xml"))
            for part_name, content_type in added_content_types.items():
                etree.SubElement(ct_root, f"{{{NS['ct']}}}Override", PartName=part_name, ContentType=content_type)
            replaced["[Content_Types].xml"] = etree.tostring(ct_root, xml_declaration=True, encoding="UTF-8",
                                                             standalone=True)

        tmp_out = dst.with_name(f".{dst.stem}.mathtype-prefs.tmp{dst.suffix}")
        with zipfile.ZipFile(tmp_out, "w", zipfile.ZIP_DEFLATED) as zout:
            names = zin.namelist()
            for name in names:
                info = zin.getinfo(name)
                zout.writestr(info, replaced.get(name, zin.read(name)))
            for name, data in replaced.items():
                if name not in names:
                    zout.writestr(name, data)
        zin.close()
        os.replace(tmp_out, dst)
        return {
            "ok": True,
            "action": "apply-equation-preferences",
            "input_path": str(src),
            "output_path": str(dst),
            "preferences": str(eqp),
            "equations_restyled": len(report),
            "document_preferences_stored": bool(store_document_preferences),
            "equations": report[:200],
        }
    finally:
        shutil.rmtree(work, ignore_errors=True)


def check_docx(document_path: str, preferences: str = "") -> dict:
    """Read-only: verify every MathType equation is typeset with the preference file.

    Each equation's MTEF is re-typeset with the preferences in memory; an equation already in that
    format keeps its size, so the rendered size must match the object's size in the document.
    """
    from lxml import etree

    path = Path(document_path).resolve()
    eqp = Path(preferences).resolve() if preferences else DEFAULT_PREFERENCES
    work = Path(tempfile.mkdtemp(prefix="mathtype-prefs-check-"))
    errors: list[str] = []
    mismatched: list[str] = []
    missing_size: list[str] = []
    unreadable: list[str] = []
    checked = 0
    try:
        zin = zipfile.ZipFile(path)
        with MathTypeApi() as api:
            prefs = api.preferences_from_file(eqp)
            for part in [n for n in zin.namelist()
                         if re.fullmatch(r"word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml", n)]:
                root = etree.fromstring(zin.read(part))
                _rels_name, rels = _rels_map(zin, part)
                for obj in root.findall(".//w:object", NS):
                    ole = obj.find("o:OLEObject", NS)
                    shape = obj.find("v:shape", NS)
                    if ole is None or shape is None or not str(ole.get("ProgID", "")).startswith("Equation.DSMT"):
                        continue
                    ole_part = "?"
                    try:
                        ole_part = _resolve(part, rels[ole.get(f"{{{NS['r']}}}id")])
                        ole_file = work / f"check-{checked}.bin"
                        ole_file.write_bytes(zin.read(ole_part))
                        _header, mtef = read_equation_native(ole_file)
                        wmf = work / f"check-{checked}.wmf"
                        api.render_wmf(api.restyle(mtef, prefs), prefs, wmf)
                        width, height = wmf_size_points(wmf.read_bytes())
                    except Exception as exc:  # one unreadable object must not abort the whole check
                        checked += 1
                        unreadable.append(f"{ole_part}: {type(exc).__name__}: {exc}")
                        continue
                    style = shape.get("style", "")
                    doc_w = _css_length_pt(style, "width")
                    doc_h = _css_length_pt(style, "height")
                    checked += 1
                    if doc_w is None or doc_h is None:
                        missing_size.append(ole_part)
                    elif abs(doc_w - width) > 0.75 or abs(doc_h - height) > 0.75:
                        mismatched.append(f"{ole_part}: {doc_w:g}x{doc_h:g} pt, expected {width:g}x{height:g} pt")
        custom = zin.read("docProps/custom.xml").decode("utf-8") if "docProps/custom.xml" in zin.namelist() else ""
        stored = 'name="MTPreferences"' in custom
        zin.close()
    finally:
        shutil.rmtree(work, ignore_errors=True)
    errors += _summarize(mismatched, f"equation(s) are not sized for {eqp.name}",
                         "run apply_mathtype_equation_preferences")
    errors += _summarize(missing_size, "equation shape(s) have no width/height in their VML style",
                         "run apply_mathtype_equation_preferences")
    errors += _summarize(unreadable, "equation object(s) could not be read", "re-insert them in MathType")
    return {"ok": not errors, "preferences": str(eqp), "equations_checked": checked,
            "document_preferences_stored": stored, "errors": errors}


def _summarize(items: list[str], what: str, fix: str, limit: int = 5) -> list[str]:
    """One summary line per problem type, listing at most ``limit`` examples, instead of one line per object."""
    if not items:
        return []
    shown = "; ".join(items[:limit]) + (f"; ... and {len(items) - limit} more" if len(items) > limit else "")
    return [f"{len(items)} {what}; {fix}. Examples: {shown}"]


def apply_to_ole_directory(directory: str, preferences: str = "") -> dict:
    """PPTX helper: restyle every oleObject*.bin in a directory, writing <name>.wmf next to it."""
    eqp = Path(preferences).resolve() if preferences else DEFAULT_PREFERENCES
    results = []
    with MathTypeApi() as api:
        prefs = api.preferences_from_file(eqp)
        for ole in sorted(Path(directory).glob("*.bin")):
            width, height = restyle_ole_file(api, prefs, ole, ole.with_suffix(".wmf"))
            results.append({"file": ole.name, "size_pt": [round(width, 2), round(height, 2)]})
    return {"ok": True, "action": "apply-equation-preferences-ole", "preferences": str(eqp), "objects": results}


def full_size_points(preferences: str = "") -> float:
    eqp = Path(preferences).resolve() if preferences else DEFAULT_PREFERENCES
    match = re.search(r"^Full=([\d.]+)\s*pt", eqp.read_text(encoding="utf-8-sig", errors="replace"), re.M)
    return float(match.group(1)) if match else 12.0



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
    d = sub.add_parser("docx")
    d.add_argument("input")
    d.add_argument("output")
    d.add_argument("--preferences", default="")
    d.add_argument("--overwrite", action="store_true")
    d.add_argument("--no-document-preferences", action="store_true")
    o = sub.add_parser("ole-dir")
    o.add_argument("directory")
    o.add_argument("--preferences", default="")
    c = sub.add_parser("check")
    c.add_argument("document")
    c.add_argument("--preferences", default="")
    args = parser.parse_args()
    try:
        if args.cmd == "docx":
            result = apply_to_docx(args.input, args.output, args.preferences, args.overwrite,
                                   not args.no_document_preferences)
        elif args.cmd == "check":
            result = check_docx(args.document, args.preferences)
        else:
            result = apply_to_ole_directory(args.directory, args.preferences)
    except Exception as exc:
        result = {"ok": False, "action": "apply-equation-preferences", "error": f"{type(exc).__name__}: {exc}"}
    sys.stdout.write(json.dumps(result, ensure_ascii=False) + "\n")
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
