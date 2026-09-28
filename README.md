# MathType Office Toolkit

[简体中文](README-zhCN.md) | [繁體中文](README-zhTW.md)

An MCP server and a cross-agent skill that let AI agents (Claude Code, Claude Desktop, Codex, ChatGPT) create **real, editable MathType 7 equations** in Microsoft Word and PowerPoint — with MathType-native equation numbers `(1)`, dynamic cross-references, and structural validation.

Version **1.0.0**. Forked from [felimet/mathtype-for-word](https://github.com/felimet/mathtype-for-word) (MIT) and extended; see [What this fork adds](#what-this-fork-adds).

---

## Contents

- [When to use it](#when-to-use-it)
- [Requirements](#requirements)
- [Core capabilities](#core-capabilities)
- [MCP tools and skill](#mcp-tools-and-skill)
- [Output format](#output-format)
- [Silent AI-agent operation](#silent-ai-agent-operation)
- [Installation](#installation)
- [Verify the installation](#verify-the-installation)
- [Troubleshooting highlights](#troubleshooting-highlights)
- [What this fork adds](#what-this-fork-adds)
- [Repository layout](#repository-layout)
- [Credits and license](#credits-and-license)

## When to use it

Use this toolkit when an AI agent has to put **mathematics into Office files that people will keep editing in MathType**:

- Writing or revising theses, journal papers, reports and course designs in Word that contain formulas.
- Converting TeX (from notes, LaTeX drafts or an LLM) into genuine MathType objects instead of Word OMath, images or Unicode text.
- Adding numbered display equations `(1), (2), …` and in-text references such as "Eq. (2)" that renumber automatically when equations are added, moved or deleted.
- Putting consistent, editable equations on PowerPoint slides for talks and defences.
- Checking that a document's equations, numbers and references are structurally correct (no `Error! Reference source not found.`, no leftover placeholders, no OMath mixed in).

It is **not** for Word's built-in equation editor (OMath), LaTeX/PDF output, macOS, or Office on the web.

## Requirements

| Item | Requirement |
|---|---|
| OS | Windows 10 or 11, interactive desktop session (Office COM automation) |
| Office | Microsoft **Word** and **PowerPoint** desktop (Microsoft 365 / Office 2016 or later; tested with 16.0) |
| MathType | Desktop **MathType for Windows** 7 from the [MathType download page](https://mathtype.tw/download/). Developed and tested with **MathType-win-zh-7.11.1.462** (`ProductVersion 7.11.1.462`). The **MathType Add-In for Microsoft 365** (task-pane add-in) is *not* enough: it lacks the desktop OLE server, Word template and PowerPoint add-in used here. |
| PowerShell | **PowerShell 7+** as `pwsh.exe`; the Office bridge never runs in Windows PowerShell 5.1 |
| Python | Python 3 on `PATH`. The MCP server uses only the standard library; the optional table layout also needs `pywin32` (`pip install -r requirements.txt`) |
| AI host | Any MCP-capable agent: Claude Code, Claude Desktop, Codex; ChatGPT through a remote endpoint or Secure MCP Tunnel |

### Terminal compatibility

The bridge is a PowerShell 7 script. The calling terminal may be PowerShell 7, Bash on Windows including Git Bash, or CMD, but the bridge itself always runs through `pwsh.exe`.

| Active terminal | Required action |
|---|---|
| PowerShell 7+ | Run the commands directly with `pwsh.exe`. |
| Bash on Windows, including Git Bash | Invoke Windows `pwsh.exe`. |
| WSL Bash | Invoke Windows `pwsh.exe`; Linux `pwsh` cannot automate Windows Office COM. |
| CMD | Invoke `pwsh.exe` with the same arguments. |
| Windows PowerShell 5.1 | Do not run the bridge in PowerShell 5.1. Switch to Git Bash or CMD and invoke `pwsh.exe`. |
| No supported terminal or no `pwsh.exe` | Install PowerShell 7; see the [Microsoft PowerShell update FAQ](https://learn.microsoft.com/zh-tw/powershell/scripting/install/microsoft-update-faq?view=powershell-7.6). |

## Core capabilities

**Word (`.docx`)**

1. **Genuine MathType objects** — every equation is an editable `Equation.DSMT4` OLE object created through MathType's own TeX conversion (`MTCommand_TeXToggle`); double-clicking opens it in MathType.
2. **Inline and display equations** — inline math stays in the sentence; display equations get their own paragraph.
3. **MathType-native numbering** — numbered displays use MathType's `MACROBUTTON MTPlaceRef` + `SEQ MTEqn` fields, formatted as simple `(1), (2), (3)` with no chapter or section part.
4. **Dynamic cross-references** — "Eq. (2)" is a MathType reference (`GOTOBUTTON` + nested `REF` to a `ZEqnNum…` bookmark), so it follows renumbering.
5. **Field update** after equations are added, moved or deleted.
6. **Optional table layout** — display equations become a 1×3 borderless table (equation centred, number right-aligned) while numbers and references stay MathType-native.
7. **Structural validation** — counts MathType objects, number and reference fields, bookmarks and sequential values; rejects OMath, leftover markers and broken references.
8. **Whole-document classification** — the skill scans the manuscript and decides which expressions are inline, unnumbered display, numbered display, or references.

**PowerPoint (`.pptx`)**

9. **Editable floating MathType equations**, horizontally centred, named `MathType_<id>`, with their embedded MathML checked against the request.
10. **Uniform equation size** — all equations use one math font size, as in Word, so a simple `σ = Eε` and a fraction look consistent.

**Safety**

11. The source file is never modified; output goes to a new path and is published atomically.
12. Word, PowerPoint and MathType run hidden and silently, and Word or PowerPoint windows the user already has open are never closed.

## MCP tools and skill

**Skill name:** `mathtype-for-word` (folder `skills/mathtype-for-word/`, packaged as `dist/mathtype-for-word.skill`). It tells the agent how to work: scan the document, classify equations, write the manifest, apply academic typography, call the tools, and validate.

**MCP server name:** `mathtype-for-word` (entry point `scripts/run-mcp.ps1` → `scripts/mcp_server.py`, stdio). Tools:

| Tool | Office | Read-only | Purpose |
|---|---|:---:|---|
| `probe_mathtype_word` | Word | ✓ | Check Windows, PowerShell, Word COM, MathType 7, its Word template and `Equation.DSMT4` registration. |
| `probe_mathtype_powerpoint` | PowerPoint | ✓ | Same checks plus PowerPoint COM and the MathType PowerPoint add-in. |
| `configure_mathtype_word_defaults` | Word | | Save the default number format `(1)` and MathType warning preferences. Run once after installing or reinstalling Office. |
| `render_mathtype_word_document` | Word | | Replace `{{MATH:id}}` / `{{EQREF:id}}` markers with MathType equations, native numbers and references (manifest-driven). |
| `apply_mathtype_repo_layout` | Word | | Optional: convert display equations to the 1×3 borderless table layout, in place. |
| `validate_mathtype_word_document` | Word | ✓ | Structural validation of objects, numbers, references, bookmarks and markers. |
| `update_mathtype_word_fields` | Word | | Refresh all number and reference fields after edits. |
| `render_mathtype_powerpoint_presentation` | PowerPoint | | Replace marker text boxes with centred, uniformly sized MathType equations. |
| `validate_mathtype_powerpoint_presentation` | PowerPoint | ✓ | Verify named objects, centring, embedded MathML, math size and leftover markers. |

A minimal Word manifest (see `examples/example-manifest.json`):

```json
{
  "schema_version": 1,
  "equations": [
    { "id": "stress", "marker": "{{MATH:stress}}", "tex": "\\sigma = \\frac{My}{I}", "layout": "display", "numbered": true }
  ],
  "references": [ { "marker": "{{EQREF:r1}}", "target": "stress" } ]
}
```

## Output format

### Word

| Element | Format |
|---|---|
| Equation object | `Equation.DSMT4` OLE, editable in MathType; never OMath, an image or text |
| Inline equation | Inline OLE object inside the sentence |
| Display equation (default) | MathType layout `<tab> equation <tab> (n)` with centre and right tab stops (`MTDisplayEquation` style) |
| Display equation (optional table layout) | 1×3 borderless table: side cells 72 pt (at most text width / 4), zero cell padding, row and lines "at least" 20 pt, vertically centred, no indents, 0 pt before and after; equation centred in the middle cell, number right-aligned in the right cell |
| Equation number | `(1), (2), (3)…` Arabic numerals in parentheses, no chapter or section, one sequence for the whole document, updated automatically. In the table layout the number uses Times New Roman / SimSun 12 pt |
| Reference | Shows `(n)` in the text; MathType `GOTOBUTTON`/`REF` field that follows renumbering |
| Typography | Scalars and variable Greek letters italic; vectors bold lowercase; matrices and tensors bold uppercase; functions, operators, constants, differentials and SI units upright; numeric sub- and superscripts upright (IEEE style, see [academic-equation-style.md](skills/mathtype-for-word/references/academic-equation-style.md)) |
| Prose | A display equation is introduced by the preceding sentence and followed by "where …" / "其中，…" defining every new symbol and unit |

### PowerPoint

| Element | Format |
|---|---|
| Equation object | Floating `Equation.DSMT4` OLE, horizontally centred, placed where the marker text box was |
| Size | One math font size for the deck, taken from `font_pt`, `equation_font_pt`, or the marker text's font size (default 24 pt); the natural 12 pt MathType object is scaled by `font_pt / 12` |
| Numbering and references | Not available: PowerPoint has no MathType field mechanism, and it is not imitated with typed numbers |

## Silent AI-agent operation

When an AI agent edits Word, PowerPoint, or MathType content, it must operate silently in the background: do not show or activate application windows, steal keyboard focus, display modal dialogs, or automate visible UI with mouse or keyboard input. If a requested step cannot be completed silently, stop and report the limitation instead of taking over the user's desktop. The bridge briefly uses the Windows clipboard to move a converted object into PowerPoint.

## Installation

### Install with an AI agent

Paste this prompt into Claude Code, Claude Desktop, Codex, or ChatGPT Desktop:

```text
Install or upgrade the MathType Office Toolkit from https://github.com/xyj0727/mathtype-office-toolkit. Detect my available terminal and use PowerShell 7, Bash on Windows including Git Bash, or CMD. Do not run the Office bridge under Windows PowerShell 5.1; if 5.1 is active, switch to Git Bash or CMD and invoke Windows pwsh.exe. From WSL Bash, invoke Windows pwsh.exe rather than Linux pwsh. If no supported terminal or pwsh.exe is available, stop and tell me to install PowerShell 7 using https://learn.microsoft.com/zh-tw/powershell/scripting/install/microsoft-update-faq?view=powershell-7.6. Verify desktop MathType for Windows ProductVersion 7.11.1.462 plus Microsoft Word and PowerPoint desktop, install the mathtype-for-word skill, register the local stdio MCP server named mathtype-for-word, run configure_mathtype_word_defaults, both MathType probes and the repository tests, preserve existing agent configuration, and report every changed file. Do not claim success unless the outputs contain editable Equation.DSMT4 objects and validation returns ok: true.
```

Platform-specific paths are in the [installation matrix](skills/mathtype-for-word/references/installation-matrix.md).

### Manual installation

1. Clone the repository and note its absolute path as `<REPO_ROOT>`:

   ```console
   git clone https://github.com/xyj0727/mathtype-office-toolkit.git
   ```

2. Optional, for the table layout: `pip install -r requirements.txt`.
3. Add the skill and register the MCP server for your agent, keeping existing MCP entries.

### Claude Code

```console
xcopy /E /I "<REPO_ROOT>\skills\mathtype-for-word" "%USERPROFILE%\.claude\skills\mathtype-for-word"
claude mcp add --scope user mathtype-for-word -- pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<REPO_ROOT>\scripts\run-mcp.ps1"
```

The repository also ships `.claude-plugin/plugin.json`, `.mcp.json` and `dist/mathtype-for-word-plugin.zip` for plugin installation.

### Claude Desktop

Upload `dist/mathtype-for-word.skill` under **Customize > Skills**, then merge this into `%APPDATA%\Claude\claude_desktop_config.json` and restart:

```json
{
  "mcpServers": {
    "mathtype-for-word": {
      "command": "pwsh.exe",
      "args": ["-NoLogo", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", "<REPO_ROOT>\\scripts\\run-mcp.ps1"]
    }
  }
}
```

### Codex

Copy `skills/mathtype-for-word` to `%USERPROFILE%\.codex\skills\mathtype-for-word`, then:

```console
codex mcp add mathtype-for-word -- pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<REPO_ROOT>\scripts\run-mcp.ps1"
```

### ChatGPT Desktop

ChatGPT cannot start a local stdio server; it needs a remote MCP endpoint or a [Secure MCP Tunnel](https://help.openai.com/en/articles/12584461-developer-mode-and-full-mcp-connectors-in-chatgpt-beta) routed to the Windows machine that runs Office.

### After installing

Run `configure_mathtype_word_defaults` once, and again after reinstalling Office.

## Verify the installation

```console
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File scripts/mathtype-word.ps1 -Action probe
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File scripts/mathtype-word.ps1 -Action probe-pptx
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File tests/run-tests.ps1 -IncludeLiveOffice
```

### Quick AI-agent test prompt

```text
Use the installed MathType Office Toolkit for a smoke test. Run both prerequisite probes, then use evals/fixtures/en-paper-draft.docx with en-word-manifest.json and evals/fixtures/en-presentation-draft.pptx with en-powerpoint-manifest.json to create new DOCX and PPTX outputs in a normal folder (not %TEMP%). Keep Word, PowerPoint, and MathType silent and hidden; do not overwrite the source fixtures. Validate both outputs and report their paths, MathType object counts, Word native number/reference counts, and the PowerPoint mathml_verified count. Do not claim success unless both validations return ok: true.
```

## Troubleshooting highlights

| Symptom | Fix |
|---|---|
| Render hangs at "Insert Equation Number" (often after reinstalling Office) | Run `configure_mathtype_word_defaults`. The MathType Word add-in reads the `HKCU\Software\Design Science\DSMT7\WordCommands` values as REG_SZ strings. |
| "property `Content` not found" right after opening the DOCX | The file is under `%TEMP%` and opens in Protected View; use a normal folder. |
| `Error! Reference source not found.` | A `ZEqnNum…` bookmark was deleted; recreate the reference through MathType. |

Full table: [troubleshooting.md](skills/mathtype-for-word/references/troubleshooting.md).

## What this fork adds

Compared with upstream [felimet/mathtype-for-word](https://github.com/felimet/mathtype-for-word) 1.3.0:

- **Table display layout (optional)** — new tool `apply_mathtype_repo_layout`, using the display-equation format of [word-mathtype-mcp](https://github.com/songsongshuo785-art/word-mathtype-mcp). The rendered paragraph is converted in place (`Range.ConvertToTable`, no clipboard), so numbers and references stay MathType-native. `validate_mathtype_word_document` accepts both layouts.
- **Uniform equation size in PowerPoint** — scaled to one math font size instead of a fixed 32 pt height; the validator detects resized equations and mixed sizes.
- **Never closes the user's PowerPoint** — PowerPoint is single-instance; rendering, validation and `probe_mathtype_powerpoint` now close only their own presentation.
- **Reinstall and localization fixes** — MathType warning preferences are written as REG_SZ (DWORD values made renders hang after an Office reinstall); presentation manifests without `height_points` no longer fail; bridge output is UTF-8 so localized error messages no longer break the MCP JSON.

## Repository layout

| Path | Purpose |
|---|---|
| `skills/mathtype-for-word/` | Cross-agent skill (`SKILL.md`), references and launcher |
| `scripts/mathtype-word.ps1` | Office automation bridge (Word, PowerPoint, MathType) |
| `scripts/mcp_server.py`, `scripts/run-mcp.ps1` | Dependency-free stdio MCP server and launcher |
| `scripts/repo_layout.py` | Table display layout (`apply_mathtype_repo_layout`) |
| `config/defaults.json` | Default Word equation-number profile |
| `config/repo_format_profile.json` | word-mathtype-mcp format profile used by the table layout |
| `examples/` | Example manifest |
| `evals/fixtures/` | Chinese and English DOCX/PPTX test inputs |
| `tests/` | Static, MCP protocol and live Office tests |
| `dist/` | `mathtype-for-word-plugin.zip` (plugin) and `mathtype-for-word.skill` (skill), with SHA-256 files |

Rebuild the plugin package with `python scripts/package_plugin.py`.

## Support

Please open an issue at [GitHub Issues](https://github.com/xyj0727/mathtype-office-toolkit/issues).

## Credits and license

[MIT](LICENSE).

- Original project: [felimet/mathtype-for-word](https://github.com/felimet/mathtype-for-word) by Jia-Ming Zhou (Felimet), MIT.
- Table-layout format profile (`config/repo_format_profile.json`): [word-mathtype-mcp](https://github.com/songsongshuo785-art/word-mathtype-mcp) by Songchongyang, MIT; see `config/LICENSE-word-mathtype-mcp.txt`.
- MathType is a trademark of its owner; this project is not affiliated with it.
