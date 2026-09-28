#!/usr/bin/env python3
"""Dependency-free stdio MCP server for native MathType automation in Word and PowerPoint."""

from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import traceback
import uuid
from pathlib import Path
from typing import Any, BinaryIO


SERVER_NAME = "mathtype-for-word"
SERVER_VERSION = "1.2.0"
SCRIPT_PATH = Path(__file__).with_name("mathtype-word.ps1")
DEFAULT_PROTOCOL_VERSION = "2025-06-18"


TOOLS: list[dict[str, Any]] = [
    {
        "name": "probe_mathtype_word",
        "description": (
            "Check Windows, Microsoft Word COM, MathType 7, the Word add-in template, "
            "and Equation.DSMT4 registration. This is read-only."
        ),
        "inputSchema": {"type": "object", "properties": {}, "additionalProperties": False},
        "annotations": {
            "title": "Probe MathType and Word",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
    {
        "name": "probe_mathtype_powerpoint",
        "description": (
            "Check PowerPoint COM, MathType 7 desktop, its PowerPoint add-in, and Equation.DSMT4 "
            "registration required for direct editable MathType OLE equations in PPTX."
        ),
        "inputSchema": {"type": "object", "properties": {}, "additionalProperties": False},
        "annotations": {
            "title": "Probe MathType and PowerPoint",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
    {
        "name": "configure_mathtype_word_defaults",
        "description": (
            "Persist the skill default: simple equation numbers (1), (2), ...; no chapter "
            "or section component; whole-document application; automatic field updates; first-number "
            "warning enabled; reference warning disabled. The bridge enforces this format on every "
            "processed document and updates the matching MathType warning preferences for the user."
        ),
        "inputSchema": {"type": "object", "properties": {}, "additionalProperties": False},
        "annotations": {
            "title": "Configure MathType for Word Defaults",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
    {
        "name": "render_mathtype_word_document",
        "description": (
            "Replace manifest markers in a DOCX with genuine Equation.DSMT4 MathType OLE equations. "
            "Numbered displays use MathType-native MTPlaceRef/SEQ fields in (1) format. References use "
            "MathType's MTReference placeholder and GOTOBUTTON/REF pipeline, never Word numbered lists. "
            "Default output follows the CJGE MathType format: equations re-typeset with "
            "config/cjge_equation_preferences.eqp (10.5 pt, TNR italic variables, Symbol italic Greek, bold "
            "italic vectors), 1x3 table display layout with (n) in Times New Roman 10.5 pt, optional "
            "punctuation between equation and number, and full-width references. Manifest options: "
            "equation_preferences (path or 'none'), display_layout ('table' or 'tab'), reference_brackets "
            "('fullwidth' or 'halfwidth'), inline_line_spacing ('at_least' default: exact-spaced paragraphs "
            "holding a taller inline equation switch to 'at least'; 'keep'), per-equation punctuation. Rendering "
            "is batched (batch_size), resumable after a crash, and can run in the background."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "input_path": {"type": "string", "description": "Absolute path to the source DOCX."},
                "output_path": {"type": "string", "description": "Absolute path for the rendered DOCX."},
                "manifest_path": {"type": "string", "description": "Absolute path to a schema v1 JSON manifest."},
                "overwrite": {"type": "boolean", "default": False},
                "batch_size": {
                    "type": "integer",
                    "default": 40,
                    "description": (
                        "Equations per isolated Word session (checkpoint after each). Large documents are "
                        "rendered in chunks so Word does not exhaust itself; 0 renders everything in one session."
                    ),
                },
                "resume": {
                    "type": "boolean",
                    "default": True,
                    "description": "Continue an interrupted or failed job for the same input, output and manifest.",
                },
                "background": {
                    "type": "boolean",
                    "default": False,
                    "description": (
                        "Start the job as a detached process and return at once; poll "
                        "get_mathtype_render_status. Recommended above about 150 equations."
                    ),
                },
                "allow_unresolved_markers": {
                    "type": "boolean",
                    "default": False,
                    "description": (
                        "Partial render: {{MATH:...}} markers that are not in this manifest may remain "
                        "(reported as warnings). Markers of this manifest must still be resolved."
                    ),
                },
            },
            "required": ["input_path", "output_path", "manifest_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Render MathType Word Document",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    },
    {
        "name": "validate_mathtype_word_document",
        "description": (
            "Open a DOCX read-only and verify genuine Equation.DSMT4 objects, simple MathType-native "
            "number fields, native references and target bookmarks, sequential numbering, resolved "
            "markers, absence of Word built-in OMath equations, and (by default) that every equation "
            "is typeset in the CJGE MathType format."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "document_path": {"type": "string", "description": "Absolute path to the DOCX."},
                "manifest_path": {"type": "string", "description": "Optional absolute path to the render manifest."},
                "equation_preferences": {
                    "type": "string",
                    "description": (
                        "Also verify that every equation is typeset with this MathType preference file "
                        "(.eqp). Default: config/cjge_equation_preferences.eqp. Pass 'none' to skip."
                    ),
                },
                "allow_unresolved_markers": {
                    "type": "boolean",
                    "default": False,
                    "description": "Report leftover markers outside the manifest as warnings instead of errors.",
                },
            },
            "required": ["document_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Validate MathType Word Document",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
    {
        "name": "update_mathtype_word_fields",
        "description": (
            "Update all Word fields in a DOCX so MathType equation numbers and references refresh. "
            "Write to a new output path, or pass overwrite=true to update the input in place."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "input_path": {"type": "string"},
                "output_path": {"type": "string"},
                "overwrite": {"type": "boolean", "default": False},
            },
            "required": ["input_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Update MathType Number and Reference Fields",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
    {
        "name": "apply_mathtype_equation_preferences",
        "description": (
            "Re-typeset every MathType equation in a DOCX with a MathType preference file (.eqp), the "
            "silent equivalent of MathType's Format Equations. Default: CJGE (config/"
            "cjge_equation_preferences.eqp: Full 10.5 pt, sub/superscripts 58 %, variables Times New "
            "Roman italic, functions upright, lower-case Greek Symbol italic, vectors/matrices bold italic). "
            "Edits the file package directly (no clipboard, no dialogs); numbers, references and fields are "
            "untouched, and the preferences are stored in the document for equations added later."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "input_path": {"type": "string", "description": "Absolute path to the DOCX."},
                "output_path": {"type": "string", "description": "Absolute path for the re-typeset DOCX."},
                "preferences_path": {"type": "string", "description": "Optional .eqp file; default CJGE."},
                "overwrite": {"type": "boolean", "default": False},
            },
            "required": ["input_path", "output_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Apply MathType Equation Preferences (CJGE)",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
    {
        "name": "apply_mathtype_repo_layout",
        "description": (
            "Convert MathType display equations to a 1x3 borderless table: equation centred, native MathType "
            "number right-aligned. Default profile CJGE (config/cjge_layout_profile.json: number Times New "
            "Roman 10.5 pt, 'at least' 15.6 pt lines, 72 pt side cells, zero padding); the word-mathtype-mcp "
            "profile is config/repo_format_profile.json. Converts in place, so MathType objects, number "
            "fields, ZEqnNum bookmarks and references are preserved. render_mathtype_word_document already "
            "applies it unless the manifest sets display_layout to 'tab'."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "input_path": {"type": "string", "description": "Absolute path to the rendered DOCX."},
                "output_path": {"type": "string", "description": "Absolute path for the formatted DOCX."},
                "profile_path": {
                    "type": "string",
                    "description": "Optional word-mathtype-mcp JSON profile; defaults to config/repo_format_profile.json.",
                },
                "overwrite": {"type": "boolean", "default": False},
            },
            "required": ["input_path", "output_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Apply word-mathtype-mcp Equation Layout",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
    {
        "name": "get_mathtype_render_status",
        "description": (
            "Progress of a batched or background render_mathtype_word_document job, identified by its output "
            "path: status (running, completed, failed, interrupted, none), equations done/total, chunks, the last "
            "log lines and the final result. Read-only unless cleanup=true removes a completed job's state."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "output_path": {"type": "string", "description": "The output_path passed to the render call."},
                "cleanup": {"type": "boolean", "default": False},
            },
            "required": ["output_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "MathType Render Job Status",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
    {
        "name": "scan_plain_text_math",
        "description": (
            "Find plain-text formulas and math symbols in a DOCX (body and table cells; the reference list is "
            "skipped) and write a reviewable candidates JSON: text, context, suggested MathType TeX (descriptive "
            "subscripts upright, units upright, Greek italic) and a proposed action. strategy 'cjge' (default) "
            "proposes MathType for expressions and italic Times New Roman text with Word subscripts for single "
            "symbols, as the CJGE guidelines require; 'all' proposes MathType for everything. Review the file "
            "(edit action to mathtype/italic_text/skip, or tex), then call prepare_mathtype_markers."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "document_path": {"type": "string", "description": "Absolute path to the DOCX."},
                "candidates_path": {"type": "string", "description": "Absolute path for the candidates JSON."},
                "strategy": {"type": "string", "enum": ["cjge", "all"], "default": "cjge"},
                "include_references": {"type": "boolean", "default": False},
                "index_subscripts": {
                    "type": "string",
                    "default": "i,j,k",
                    "description": "Comma-separated subscript letters that are indices (kept italic).",
                },
            },
            "required": ["document_path", "candidates_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Scan Plain-Text Math",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
    {
        "name": "prepare_mathtype_markers",
        "description": (
            "Apply a reviewed candidates JSON from scan_plain_text_math to a copy of the DOCX: 'mathtype' "
            "candidates become unique {{MATH:...}} markers (across runs, in tables) and a schema v1 manifest is "
            "written; 'italic_text' candidates become Times New Roman italic text with real subscripts; 'skip' is "
            "untouched. Then call render_mathtype_word_document with the new DOCX and manifest."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "document_path": {"type": "string", "description": "The scanned source DOCX (unchanged since the scan)."},
                "candidates_path": {"type": "string"},
                "output_path": {"type": "string", "description": "Absolute path for the marked DOCX."},
                "manifest_path": {"type": "string", "description": "Absolute path for the generated manifest."},
                "overwrite": {"type": "boolean", "default": False},
                "id_prefix": {"type": "string", "default": "eq"},
                "index_subscripts": {"type": "string", "default": "i,j,k"},
            },
            "required": ["document_path", "candidates_path", "output_path", "manifest_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Prepare MathType Markers",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    },
    {
        "name": "fix_mathtype_line_spacing",
        "description": (
            "Switch paragraphs (and table-cell paragraphs) whose exact line spacing is lower than an inline "
            "MathType equation to 'at least' with the same value, so equations are not clipped. The spacing may "
            "come from the paragraph, its style chain or docDefaults. render_mathtype_word_document does this "
            "automatically unless the manifest sets inline_line_spacing to 'keep'."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "input_path": {"type": "string"},
                "output_path": {"type": "string"},
                "overwrite": {"type": "boolean", "default": False},
            },
            "required": ["input_path", "output_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Fix Line Spacing Around Inline MathType",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
    {
        "name": "apply_cjge_body_format",
        "description": (
            "Format the non-equation parts of a DOCX as a CJGE manuscript from config/cjge_body_profile.json: "
            "A4 page and margins; title, heading 1-3, body, list, table and reference fonts, sizes, line spacing "
            "and indents (SimSun/SimHei/KaiTi + Times New Roman); three-line tables. MathType objects, equation "
            "tables and runs holding equations are left untouched. Writes a new DOCX."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "input_path": {"type": "string"},
                "output_path": {"type": "string"},
                "profile_path": {"type": "string", "description": "Optional body profile JSON; default CJGE."},
                "overwrite": {"type": "boolean", "default": False},
            },
            "required": ["input_path", "output_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Apply CJGE Body Format",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
    {
        "name": "report_docx_formatting",
        "description": (
            "Compact, read-only formatting summary of a DOCX for review instead of dumping raw XML: page size and "
            "margins; per paragraph role (title, heading levels, body, list, table, reference) the effective East "
            "Asian/Latin fonts, sizes, bold, line spacing, indents and alignment with counts and examples; table "
            "border styles; MathType object count. With a profile it also lists deviations from it."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "document_path": {"type": "string"},
                "profile_path": {
                    "type": "string",
                    "description": "Optional body profile to compare against; 'cjge' uses config/cjge_body_profile.json.",
                },
            },
            "required": ["document_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Report DOCX Formatting",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
    {
        "name": "render_mathtype_powerpoint_presentation",
        "description": (
            "Silently replace marker-only PPTX text boxes with editable, centered Equation.DSMT4 floating OLE "
            "objects through hidden Word MathType conversion and MathML validation. PowerPoint has no "
            "MathType-native Word equation numbering/reference mechanism."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "input_path": {"type": "string", "description": "Absolute path to the source PPTX."},
                "output_path": {"type": "string", "description": "Absolute path for the rendered PPTX."},
                "manifest_path": {"type": "string", "description": "Absolute path to a presentation schema v1 manifest."},
                "overwrite": {"type": "boolean", "default": False},
            },
            "required": ["input_path", "output_path", "manifest_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Render MathType PowerPoint Presentation",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    },
    {
        "name": "validate_mathtype_powerpoint_presentation",
        "description": (
            "Verify that a PPTX contains the expected named, centered Equation.DSMT4 OLE objects, "
            "that their embedded MathML matches the manifest, and that no markers remain."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "presentation_path": {"type": "string", "description": "Absolute path to the PPTX."},
                "manifest_path": {"type": "string", "description": "Absolute path to the presentation manifest."},
            },
            "required": ["presentation_path", "manifest_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Validate MathType PowerPoint Presentation",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
]


def _log(message: str) -> None:
    try:
        print(f"[{SERVER_NAME}] {message}", file=sys.stderr, flush=True)
    except (UnicodeError, OSError):
        pass  # logging must never break a tool call


def _utf8_stdio() -> None:
    """Windows starts Python with an ANSI stdio code page; CJK paths in results and logs need UTF-8."""
    for stream in (sys.stdin, sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass


def _powershell_executable() -> str:
    return os.environ.get("MATHTYPE_WORD_POWERSHELL", "pwsh.exe")


def _warning_preferences() -> dict[str, int]:
    try:
        import winreg

        path = r"Software\Design Science\DSMT7\WordCommands"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:
            return {
                name: int(winreg.QueryValueEx(key, name)[0])
                for name in ("NoEqnNumWarningDlg", "NoInsertEqnRefDlg")
            }
    except (ImportError, OSError, ValueError):
        return {}


def _restore_warning_preferences(preferences: dict[str, int]) -> None:
    if not preferences:
        return
    try:
        import winreg

        path = r"Software\Design Science\DSMT7\WordCommands"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_SET_VALUE) as key:
            for name, value in preferences.items():
                # The MathType Word add-in reads these preferences as REG_SZ strings ("0"/"1").
                winreg.SetValueEx(key, name, 0, winreg.REG_SZ, str(value))
    except (ImportError, OSError, ValueError) as exc:
        _log(f"could not restore MathType warning preferences after timeout: {exc}")


def _extract_word_pid(stderr: str, key: str = "WORD_PID") -> int | None:
    matches = re.findall(rf"\b{key}=(\d+)\b", stderr)
    return int(matches[-1]) if matches else None


def _process_pids(executable_name: str) -> set[int]:
    if os.name != "nt":
        return set()
    try:
        import ctypes
        from ctypes import wintypes

        process_ids = (wintypes.DWORD * 4096)()
        bytes_returned = wintypes.DWORD()
        if not ctypes.windll.psapi.EnumProcesses(
            ctypes.byref(process_ids), ctypes.sizeof(process_ids), ctypes.byref(bytes_returned)
        ):
            return set()
        count = bytes_returned.value // ctypes.sizeof(wintypes.DWORD)
        matches: set[int] = set()
        for process_id in process_ids[:count]:
            handle = ctypes.windll.kernel32.OpenProcess(0x1000, False, process_id)
            if not handle:
                continue
            try:
                capacity = wintypes.DWORD(32768)
                buffer = ctypes.create_unicode_buffer(capacity.value)
                if ctypes.windll.kernel32.QueryFullProcessImageNameW(
                    handle, 0, buffer, ctypes.byref(capacity)
                ) and Path(buffer.value).name.upper() == executable_name.upper():
                    matches.add(int(process_id))
            finally:
                ctypes.windll.kernel32.CloseHandle(handle)
        return matches
    except (AttributeError, OSError, ValueError):
        return set()


def _terminate_process_pid(process_id: int) -> bool:
    completed = subprocess.run(
        ["taskkill.exe", "/PID", str(process_id), "/T", "/F"],
        stdin=subprocess.DEVNULL,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    return completed.returncode == 0


def _cleanup_bridge_output(output_path: str | None, run_token: str) -> dict[str, Any]:
    if not output_path:
        return {"ok": True, "skipped": "No output path was supplied."}
    command = [
        _powershell_executable(),
        "-NoLogo",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(SCRIPT_PATH),
        "-Action",
        "cleanup",
        "-OutputPath",
        output_path,
        "-RunToken",
        run_token,
    ]
    try:
        completed = subprocess.run(
            command,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=30,
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        return {"ok": False, "error": f"Cleanup bridge failed: {exc}"}
    lines = [line.strip() for line in completed.stdout.splitlines() if line.strip()]
    if not lines:
        return {"ok": False, "error": f"Cleanup bridge returned no JSON (exit {completed.returncode})."}
    try:
        return json.loads(lines[-1])
    except json.JSONDecodeError as exc:
        return {"ok": False, "error": f"Cleanup bridge returned invalid JSON: {exc}"}


def _invoke_bridge(action: str, arguments: dict[str, Any], timeout: int | None = None) -> dict[str, Any]:
    run_token = uuid.uuid4().hex
    command = [
        _powershell_executable(),
        "-NoLogo",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(SCRIPT_PATH),
        "-Action",
        action,
        "-RunToken",
        run_token,
    ]
    mapping = {
        "input_path": "-InputPath",
        "document_path": "-InputPath",
        "presentation_path": "-InputPath",
        "output_path": "-OutputPath",
        "manifest_path": "-ManifestPath",
    }
    for key, switch in mapping.items():
        value = arguments.get(key)
        if value is not None and value != "":
            command.extend([switch, str(value)])
    if arguments.get("overwrite"):
        command.append("-Overwrite")
    if arguments.get("allow_unresolved_markers"):
        command.append("-AllowUnresolvedMarkers")

    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    effective_timeout = timeout or int(
        os.environ.get("MATHTYPE_OFFICE_TIMEOUT_SECONDS", os.environ.get("MATHTYPE_WORD_TIMEOUT_SECONDS", "240"))
    )
    preferences = _warning_preferences()
    word_pids_before = _process_pids("WINWORD.EXE")
    power_point_pids_before = _process_pids("POWERPNT.EXE")
    try:
        completed = subprocess.run(
            command,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=effective_timeout,
            check=False,
            creationflags=creationflags,
        )
    except subprocess.TimeoutExpired as exc:
        stderr = exc.stderr or ""
        if isinstance(stderr, bytes):
            stderr = stderr.decode("utf-8", errors="replace")
        word_process_id = _extract_word_pid(stderr)
        power_point_process_id = _extract_word_pid(stderr, "POWERPOINT_PID")
        new_word_pids = sorted(_process_pids("WINWORD.EXE") - word_pids_before)
        new_power_point_pids = sorted(_process_pids("POWERPNT.EXE") - power_point_pids_before)
        if word_process_id is None and len(new_word_pids) == 1:
            word_process_id = new_word_pids[0]
        if power_point_process_id is None and len(new_power_point_pids) == 1:
            power_point_process_id = new_power_point_pids[0]
        # Terminate only the Office processes this bridge started (logged PID, or the single new one);
        # Word or PowerPoint windows the user opened meanwhile are never touched.
        targets = [pid for pid in (word_process_id, power_point_process_id) if pid]
        terminated = [process_id for process_id in targets if _terminate_process_pid(process_id)]
        untouched = sorted(set(new_word_pids + new_power_point_pids) - set(targets))
        if untouched:
            _log(f"left unidentified new Office processes running: {untouched}")
        cleaned = bool(word_process_id and word_process_id in terminated)
        _restore_warning_preferences(preferences)
        cleanup = _cleanup_bridge_output(arguments.get("output_path"), run_token)
        return {
            "ok": False,
            "action": action,
            "error": f"Word/MathType bridge timed out after {effective_timeout} seconds.",
            "isolated_word_pid": word_process_id,
            "isolated_word_process_terminated": cleaned,
            "new_word_pids_observed": new_word_pids,
            "new_powerpoint_pids_observed": new_power_point_pids,
            "isolated_office_processes_terminated": terminated,
            "unidentified_office_processes_left_running": untouched,
            "temporary_file_cleanup": cleanup,
        }
    if completed.stderr.strip():
        _log(completed.stderr.strip())
    lines = [line.strip() for line in completed.stdout.splitlines() if line.strip()]
    if not lines:
        return {
            "ok": False,
            "action": action,
            "error": f"Bridge returned no JSON (exit {completed.returncode}).",
        }
    try:
        result = json.loads(lines[-1])
    except json.JSONDecodeError as exc:
        return {
            "ok": False,
            "action": action,
            "error": f"Bridge returned invalid JSON: {exc}",
            "stdout": completed.stdout[-4000:],
        }
    if completed.returncode and result.get("ok", True):
        result["ok"] = False
        result["error"] = f"Bridge exited with code {completed.returncode}."
    return result


def _apply_repo_layout(arguments: dict[str, Any]) -> dict[str, Any]:
    command = [
        sys.executable,
        str(Path(__file__).with_name("repo_layout.py")),
        str(arguments["input_path"]),
        str(arguments["output_path"]),
    ]
    if arguments.get("profile_path"):
        command += ["--profile", str(arguments["profile_path"])]
    if arguments.get("overwrite"):
        command.append("--overwrite")
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    try:
        completed = subprocess.run(
            command, capture_output=True, text=True, encoding="utf-8", timeout=600, env=env, check=False
        )
    except subprocess.TimeoutExpired:
        return {"ok": False, "action": "apply-repo-layout", "error": "Timed out after 600 s."}
    lines = [line for line in completed.stdout.splitlines() if line.strip()]
    try:
        return json.loads(lines[-1])
    except (IndexError, json.JSONDecodeError):
        return {
            "ok": False,
            "action": "apply-repo-layout",
            "error": "repo_layout.py returned no JSON result.",
            "stderr": completed.stderr[-2000:],
        }


def _run_python_tool(script: str, argv: list[str], action: str, timeout: int = 600) -> dict[str, Any]:
    command = [sys.executable, str(Path(__file__).with_name(script)), *argv]
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    try:
        completed = subprocess.run(
            command, capture_output=True, text=True, encoding="utf-8", timeout=timeout, env=env, check=False
        )
    except subprocess.TimeoutExpired:
        return {"ok": False, "action": action, "error": f"Timed out after {timeout} s."}
    lines = [line for line in completed.stdout.splitlines() if line.strip()]
    try:
        return json.loads(lines[-1])
    except (IndexError, json.JSONDecodeError):
        return {"ok": False, "action": action, "error": f"{script} returned no JSON result.",
                "stderr": completed.stderr[-2000:]}


def _apply_equation_preferences(arguments: dict[str, Any]) -> dict[str, Any]:
    argv = ["docx", str(arguments["input_path"]), str(arguments["output_path"])]
    if arguments.get("preferences_path"):
        argv += ["--preferences", str(arguments["preferences_path"])]
    if arguments.get("overwrite"):
        argv.append("--overwrite")
    return _run_python_tool("mathtype_prefs.py", argv, "apply-equation-preferences")


def _check_equation_preferences(document_path: str, preferences: str) -> dict[str, Any]:
    argv = ["check", document_path]
    if preferences:
        argv += ["--preferences", preferences]
    return _run_python_tool("mathtype_prefs.py", argv, "check-equation-preferences")


def _validate_word(arguments: dict[str, Any]) -> dict[str, Any]:
    result = _invoke_bridge("validate", arguments)
    if result.get("counts"):
        spacing = _run_python_tool("docx_postprocess.py", ["check", str(arguments["document_path"])], "check-line-spacing")
        result["warnings"] = list(result.get("warnings") or []) + list(spacing.get("warnings") or [])
    preferences = str(arguments.get("equation_preferences") or "")
    if preferences.lower() == "none" or not result.get("counts"):
        return result
    check = _check_equation_preferences(str(arguments["document_path"]), preferences)
    result["equation_format"] = check
    if not check.get("ok"):
        result["ok"] = False
        result["errors"] = list(result.get("errors") or []) + list(
            check.get("errors") or [check.get("error", "equation format check failed")]
        )
    return result


def _postprocess_word(
    output: str, manifest: dict[str, Any], manifest_path: str, allow_unresolved_markers: bool = False
) -> dict[str, Any]:
    """Run once after all equations exist: preferences, table layout, line spacing, strict validation."""
    preferences = str(manifest.get("equation_preferences") or "")
    layout = str(manifest.get("display_layout") or "table")
    spacing = str(manifest.get("inline_line_spacing") or "at_least")
    if layout not in ("table", "tab"):
        return {"ok": False, "error": "display_layout must be 'table' or 'tab'."}
    if spacing not in ("at_least", "keep"):
        return {"ok": False, "error": "inline_line_spacing must be 'at_least' or 'keep'."}
    steps: dict[str, Any] = {}
    if preferences.lower() != "none":
        step = _apply_equation_preferences(
            {"input_path": output, "output_path": output, "preferences_path": preferences, "overwrite": True}
        )
        steps["equation_preferences"] = step
        if not step.get("ok"):
            return {"ok": False, "error": f"Equation preferences failed: {step.get('error')}", "post_processing": steps}
    if layout == "table":
        step = _apply_repo_layout({"input_path": output, "output_path": output, "overwrite": True})
        steps["display_layout"] = step
        if not step.get("ok"):
            return {"ok": False, "error": f"Table layout failed: {step.get('error')}", "post_processing": steps}
    if spacing == "at_least":
        step = _run_python_tool("docx_postprocess.py", ["fix", output, output, "--overwrite"], "fix-line-spacing")
        steps["inline_line_spacing"] = step
        if not step.get("ok"):
            return {"ok": False, "error": f"Line-spacing fix failed: {step.get('error')}", "post_processing": steps}
    validation = _validate_word(
        {
            "document_path": output,
            "manifest_path": manifest_path,
            "equation_preferences": preferences,
            "allow_unresolved_markers": allow_unresolved_markers,
        }
    )
    return {"ok": bool(validation.get("ok")), "post_processing": steps, "validation": validation}


def _render_word(arguments: dict[str, Any]) -> dict[str, Any]:
    import mathtype_batch

    batch_size = int(arguments.get("batch_size", mathtype_batch.DEFAULT_BATCH_SIZE))
    common = (
        str(arguments["input_path"]),
        str(arguments["output_path"]),
        str(arguments["manifest_path"]),
        batch_size,
        bool(arguments.get("resume", True)),
        bool(arguments.get("overwrite")),
        bool(arguments.get("allow_unresolved_markers")),
    )
    try:
        if arguments.get("background"):
            mathtype_batch.load_manifest(common[2])
            return mathtype_batch.start_background(*common)
        return mathtype_batch.render(*common)
    except Exception as exc:  # report as a tool result, not a protocol error
        return {"ok": False, "action": "render", "error": f"{type(exc).__name__}: {exc}"}


def _scan_math(arguments: dict[str, Any]) -> dict[str, Any]:
    argv = ["scan", str(arguments["document_path"]), "--candidates", str(arguments["candidates_path"]),
            "--strategy", str(arguments.get("strategy") or "cjge"),
            "--index-subscripts", str(arguments.get("index_subscripts") or "i,j,k")]
    if arguments.get("include_references"):
        argv.append("--include-references")
    return _run_python_tool("mathtype_scan.py", argv, "scan")


def _prepare_markers(arguments: dict[str, Any]) -> dict[str, Any]:
    argv = ["prepare", str(arguments["document_path"]), str(arguments["candidates_path"]),
            str(arguments["output_path"]), str(arguments["manifest_path"]),
            "--id-prefix", str(arguments.get("id_prefix") or "eq"),
            "--index-subscripts", str(arguments.get("index_subscripts") or "i,j,k")]
    if arguments.get("overwrite"):
        argv.append("--overwrite")
    return _run_python_tool("mathtype_scan.py", argv, "prepare")


def _fix_line_spacing(arguments: dict[str, Any]) -> dict[str, Any]:
    argv = ["fix", str(arguments["input_path"]), str(arguments["output_path"])]
    if arguments.get("overwrite"):
        argv.append("--overwrite")
    return _run_python_tool("docx_postprocess.py", argv, "fix-line-spacing")


def _body_format(arguments: dict[str, Any]) -> dict[str, Any]:
    argv = ["apply", str(arguments["input_path"]), str(arguments["output_path"])]
    if arguments.get("profile_path"):
        argv += ["--profile", str(arguments["profile_path"])]
    if arguments.get("overwrite"):
        argv.append("--overwrite")
    return _run_python_tool("cjge_body_format.py", argv, "apply-cjge-body-format")


def _format_report(arguments: dict[str, Any]) -> dict[str, Any]:
    argv = ["report", str(arguments["document_path"])]
    if arguments.get("profile_path"):
        argv += ["--profile", str(arguments["profile_path"])]
    return _run_python_tool("cjge_body_format.py", argv, "report-docx-formatting")


def _render_status(arguments: dict[str, Any]) -> dict[str, Any]:
    import mathtype_batch

    return mathtype_batch.status(str(arguments["output_path"]), bool(arguments.get("cleanup")))


def _call_tool(name: str, arguments: dict[str, Any]) -> dict[str, Any]:
    actions = {
        "probe_mathtype_word": "probe",
        "probe_mathtype_powerpoint": "probe-pptx",
        "configure_mathtype_word_defaults": "configure-defaults",
        "render_mathtype_word_document": "render",
        "validate_mathtype_word_document": "validate",
        "update_mathtype_word_fields": "update",
        "render_mathtype_powerpoint_presentation": "render-pptx",
        "validate_mathtype_powerpoint_presentation": "validate-pptx",
    }
    local = {
        "get_mathtype_render_status": _render_status,
        "scan_plain_text_math": _scan_math,
        "prepare_mathtype_markers": _prepare_markers,
        "fix_mathtype_line_spacing": _fix_line_spacing,
        "apply_cjge_body_format": _body_format,
        "report_docx_formatting": _format_report,
    }
    if name in local:
        result = local[name](arguments)
    elif name == "apply_mathtype_repo_layout":
        result = _apply_repo_layout(arguments)
    elif name == "apply_mathtype_equation_preferences":
        result = _apply_equation_preferences(arguments)
    elif name == "render_mathtype_word_document":
        result = _render_word(arguments)
    elif name == "validate_mathtype_word_document":
        result = _validate_word(arguments)
    elif name not in actions:
        raise ValueError(f"Unknown tool: {name}")
    else:
        result = _invoke_bridge(actions[name], arguments)
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    return {
        "content": [{"type": "text", "text": rendered}],
        "structuredContent": result,
        "isError": not bool(result.get("ok")),
    }


def _read_message(stream: BinaryIO) -> dict[str, Any] | None:
    line = stream.readline()
    if not line:
        return None
    while line in (b"\r\n", b"\n"):
        line = stream.readline()
        if not line:
            return None
    if line.lower().startswith(b"content-length:"):
        length = int(line.split(b":", 1)[1].strip())
        while True:
            header = stream.readline()
            if header in (b"\r\n", b"\n", b""):
                break
        payload = stream.read(length)
        return json.loads(payload.decode("utf-8"))
    return json.loads(line.decode("utf-8"))


def _write_message(message: dict[str, Any]) -> None:
    data = json.dumps(message, ensure_ascii=False, separators=(",", ":"))
    sys.stdout.write(data + "\n")
    sys.stdout.flush()


def _success(request_id: Any, result: Any) -> dict[str, Any]:
    return {"jsonrpc": "2.0", "id": request_id, "result": result}


def _error(request_id: Any, code: int, message: str, data: Any = None) -> dict[str, Any]:
    error: dict[str, Any] = {"code": code, "message": message}
    if data is not None:
        error["data"] = data
    return {"jsonrpc": "2.0", "id": request_id, "error": error}


def _handle(message: dict[str, Any]) -> dict[str, Any] | None:
    method = message.get("method")
    request_id = message.get("id")
    params = message.get("params") or {}
    if method in ("notifications/initialized", "notifications/cancelled"):
        return None
    if method == "initialize":
        requested = params.get("protocolVersion") or DEFAULT_PROTOCOL_VERSION
        return _success(
            request_id,
            {
                "protocolVersion": requested,
                "capabilities": {"tools": {"listChanged": False}},
                "serverInfo": {"name": SERVER_NAME, "version": SERVER_VERSION},
                "instructions": (
                    "Use the matching Word or PowerPoint probe first. Preserve the source, render from a "
                    "schema v1 manifest, then validate. For documents with plain-text formulas, run "
                    "scan_plain_text_math, review the candidates, then prepare_mathtype_markers. Large documents "
                    "render in resumable batches; use background=true and get_mathtype_render_status for long "
                    "jobs. DOCX numbering and references must remain MathType-native fields. PPTX equations are "
                    "editable floating Equation.DSMT4 OLE objects."
                ),
            },
        )
    if method == "ping":
        return _success(request_id, {})
    if method == "tools/list":
        return _success(request_id, {"tools": TOOLS})
    if method == "tools/call":
        return _success(request_id, _call_tool(str(params.get("name", "")), params.get("arguments") or {}))
    if request_id is None:
        return None
    return _error(request_id, -32601, f"Method not found: {method}")


def main() -> int:
    _utf8_stdio()
    _log(f"starting {SERVER_VERSION}")
    while True:
        message: dict[str, Any] | None = None
        try:
            message = _read_message(sys.stdin.buffer)
            if message is None:
                return 0
            response = _handle(message)
            if response is not None:
                _write_message(response)
        except json.JSONDecodeError as exc:
            _write_message(_error(None, -32700, "Parse error", str(exc)))
        except Exception as exc:  # MCP boundary: convert failures to JSON-RPC errors.
            _log(traceback.format_exc())
            request_id = message.get("id") if isinstance(message, dict) else None
            _write_message(_error(request_id, -32603, str(exc)))


if __name__ == "__main__":
    raise SystemExit(main())
