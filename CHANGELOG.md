# Changelog

## 1.2.0 — 2026-09-28

Large documents, plain-text formulas and body-text formatting. Every item below comes from problems met while converting a real 330-equation geotechnical manuscript.

### Added

- **`scan_plain_text_math`** — finds typed formulas and math symbols in body text and table cells (the reference list is skipped) and writes a reviewable candidates JSON with context, suggested MathType TeX (descriptive subscripts, units and functions upright; Greek italic; index subscripts `i, j, k` italic) and a proposed action. Strategy `cjge` (default) proposes MathType for expressions and Times New Roman italic text with Word subscripts for single symbols, as the CJGE guidelines require; `all` proposes MathType for everything.
- **`prepare_mathtype_markers`** — applies the reviewed candidates: unique `{{MATH:eqNNN}}` markers (also across runs and in tables) plus a schema v1 manifest, italic text for `italic_text` candidates; refuses a DOCX that changed after the scan.
- **Batched, resumable rendering** — `render_mathtype_word_document` renders in batches (`batch_size`, default 40) in fresh isolated Word sessions with a checkpoint after each batch. A failed batch is retried once split in half; calling render again with the same arguments resumes (`resume`, default true). Referenced numbered equations and all references are rendered together in the last batch. Preferences, table layout, line-spacing fix and the strict validation run once at the end. Job state lives in `%APPDATA%\MathTypeForWordAgent\jobs`.
- **Background jobs** — `background: true` starts a detached job and returns at once; **`get_mathtype_render_status`** reports status, equations done/total, batches, log tail and the final result.
- **Partial renders** — `allow_unresolved_markers` (render and validate; bridge `-AllowUnresolvedMarkers`) reports markers outside the manifest as warnings.
- **`fix_mathtype_line_spacing`** and manifest field `inline_line_spacing` (`at_least` default, `keep`) — exactly spaced paragraphs that would clip an inline equation switch to "at least" with the same value (paragraph, style chain or docDefaults spacing); render applies it automatically and validation warns about clipped equations.
- **`apply_cjge_body_format`** with `config/cjge_body_profile.json` — CJGE layout for everything except equations: page and margins, title, headings 1–3, body, lists, captions, three-line tables, reference list; MathType runs and equation tables are never touched; OOXML child order is preserved.
- **`report_docx_formatting`** — compact per-role summary of effective fonts, sizes, bold, line spacing, indents, alignment and table borders, with deviations from a profile (`profile_path: "cjge"`), instead of dumping raw XML.
- CLI helpers `mathtype_scan.py`, `mathtype_batch.py`, `docx_postprocess.py`, `cjge_body_format.py` for use without MCP.
- `tests/test_v12_features.py` — 22 unit tests (scan false positives, TeX conversion, marker preparation, batching/split/resume, line spacing, body format and schema order, CSS units, summaries, UTF-8).
- `tests/test_live_regressions.py` — live Word/MathType regressions (table-cell markers, merged cells, Chinese paths, 124 equations in 3 batches, partial render, `1in` sizes), run with `tests/run-tests.ps1 -IncludeLiveOffice`. Verified on a real 321-equation manuscript: 9 batches in the background, no retries, final validation ok.

### Fixed

- Render hung forever on a marker inside a table cell: Word's `Find` returned the same table match again; the marker search now stops when a match does not move forward.
- The equation-format check crashed with `AttributeError` when Word wrote a VML size as `1in` (exactly 72 pt); all CSS length units are now read and written, and a missing size is reported instead of crashing.
- `'charmap' codec can't encode characters` when paths or messages contain Chinese: the MCP server and every Python helper now use UTF-8 stdio.
- A single Word session exhausted itself (RPC `0x800706BA`/`0x800706BE`) after roughly 200 equations — solved by batching.

### Changed

- Bridge log lines carry timestamps and `PROGRESS i/N` lines.
- On a timeout the watchdog terminates only the Office process the bridge logged as its own (or the single new one) and leaves any other new Word/PowerPoint window running.
- Repeated equation-format errors are summarised in one line per problem type with up to five examples; one unreadable equation no longer aborts the whole check.
- Server, plugin and skill version 1.2.0; skill workflow, references and READMEs updated.

## 1.1.0

CJGE equation format: `apply_mathtype_equation_preferences`, CJGE preference and layout profiles, punctuation between equation and number, 式（n） references, equation-format validation, CJGE styling in PowerPoint.

## 1.0.0

Table display layout, uniform PowerPoint equation size, the user's PowerPoint is never closed, REG_SZ warning preferences, UTF-8 bridge output.
