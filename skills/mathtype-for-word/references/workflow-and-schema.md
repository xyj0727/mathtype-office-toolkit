# Workflow and manifest schema

## Contents

1. End-to-end flow
2. Schema v1
3. Complete example
4. Marker preparation
5. Output contract

## End-to-end flow

The bridge performs a deterministic transaction:

1. Open the input DOCX in an isolated hidden Word instance.
2. Require every marker to occur exactly once.
3. Replace equation markers with TeX delimited for MathType Toggle TeX.
4. Require exactly one new `Equation.DSMT4` object per equation.
5. For numbered displays, call MathType's native number command and normalize its native fields to `(n)`.
6. Replace reference markers through MathType's `MTReference` and `MTPlaceRef` workflow.
7. Update Word fields.
8. Save to a temporary sibling file and atomically move it to the output path.
9. Reopen read-only and validate.

The bridge refuses an existing output unless `overwrite` is explicitly true.

## Schema v1

```json
{
  "schema_version": 1,
  "equations": [
    {
      "id": "stable_ascii_identifier",
      "marker": "{{MATH:stable_ascii_identifier}}",
      "tex": "one-line MathType-compatible TeX",
      "layout": "inline or display",
      "numbered": false
    }
  ],
  "references": [
    {
      "marker": "{{EQREF:stable_ascii_identifier}}",
      "target": "stable_ascii_identifier"
    }
  ]
}
```

Rules:

- `schema_version` must equal `1`.
- `equations` is required. `references` may be empty.
- `id` values are unique and case-sensitive.
- All equation and reference markers are unique across the manifest.
- A marker occurs exactly once in the DOCX.
- `tex` is non-empty and contains no newline.
- `layout` is `inline` or `display`.
- `numbered: true` is legal only for `display`.
- Optional `punctuation` (display equations only): `，` `。` `；` or `,` `.` `;`, inserted between the equation and its number (CJGE: `，` before 式中, `。` at a sentence end).

Optional top-level fields (defaults give the CJGE format, see [cjge-format.md](cjge-format.md)):

| Field | Default | Meaning |
|---|---|---|
| `reference_brackets` | `"fullwidth"` | `fullwidth` renders 式（n）; `halfwidth` keeps MathType's `(n)` |
| `equation_preferences` | CJGE `.eqp` | MathType preference file used to re-typeset every equation; `"none"` keeps MathType's current defaults |
| `display_layout` | `"table"` | `table` = 1x3 borderless table (CJGE layout profile); `tab` = MathType's tab layout |
| `inline_line_spacing` | `"at_least"` | exactly spaced paragraphs holding a taller inline equation switch to "at least" with the same value; `keep` leaves spacing alone |

Unknown equation fields such as `source_text` (written by `prepare_mathtype_markers`) are ignored.

Render call options (tool arguments, not manifest fields):

| Argument | Default | Meaning |
|---|---|---|
| `batch_size` | `40` | equations per isolated Word session; a checkpoint is saved after each batch; `0` = one session |
| `resume` | `true` | continue a failed or interrupted job for the same input, output and manifest |
| `background` | `false` | start a detached job and return at once; poll `get_mathtype_render_status` |
| `allow_unresolved_markers` | `false` | partial render: markers outside this manifest may remain (warnings) |

Numbered equations that are referenced, and all references, are rendered together in the last batch, because MathType places a reference by clicking the live number field.
- A reference target names an existing numbered equation.

Semantic classes map to schema v1 as follows:

| Class | `layout` | `numbered` | Additional entry |
|---|---|---:|---|
| `inline` | `inline` | `false` | None |
| `display` | `display` | `false` | None |
| `display_numbered` | `display` | `true` | None |
| `reference` | Not applicable | Not applicable | One unique `references[]` item targeting a `display_numbered` equation |

Read the whole document before finalizing this table. A later semantic reference promotes its target from `display` to `display_numbered`. Multiple reference occurrences may share a target but must use different source markers.

## Complete example

Source text:

```text
The relativistic energy relation is expressed as follows:

{{MATH:mass_energy}}

where E is the energy in joules, m is the mass in kilograms, and c is the speed of light in metres per second.

As shown in Eq. {{EQREF:mass_energy_result}}, energy is proportional to mass.
```

Manifest:

```json
{
  "schema_version": 1,
  "equations": [
    {
      "id": "mass_energy",
      "marker": "{{MATH:mass_energy}}",
      "tex": "E = mc^2",
      "layout": "display",
      "numbered": true
    }
  ],
  "references": [
    {
      "marker": "{{EQREF:mass_energy_result}}",
      "target": "mass_energy"
    }
  ]
}
```

Expected output structure:

- One `Equation.DSMT4` OLE object.
- One native MathType number displayed as `(1)`.
- One native reference displayed as `(1)`.
- The target number is bookmarked as `ZEqnNum` followed by digits.
- The reference is a `GOTOBUTTON` field with a nested `REF` field.

## Marker preparation

For a document with plain-text formulas, let the tools place the markers:

1. `scan_plain_text_math` writes a candidates file (`mathtype-for-word-candidates/1`): `document_sha256`, `strategy`, and per candidate `id`, `paragraph_index`, `start`, `end`, `text`, `context`, `in_table`, `kind` (`expression` or `symbol`), `tex` and `action` (`mathtype`, `italic_text`, `skip`).
2. Review and edit `action` and `tex`. Do not change `paragraph_index`, `start`, `end` or `text`.
3. `prepare_mathtype_markers` refuses a DOCX that changed after the scan, replaces `mathtype` candidates with `{{MATH:eq001}}`… (also across runs and in table cells), writes italic text with Word subscripts for `italic_text`, and writes the manifest (inline, unnumbered). Promote display equations and add numbers and references in the manifest by hand when needed.

For markers you place yourself, prefer markers that are visually obvious, unique, and absent from normal prose. Put a display marker alone in its paragraph. An inline marker remains inside the sentence.

Good:

```text
The gain is {{MATH:inline_gain}} at resonance.

{{MATH:transfer_function}}

其中，H(s) 為轉移函數，s 為複數頻率，單位為 rad/s。
```

Avoid:

- Reusing `{{MATH:eq1}}` twice.
- Embedding a display marker inside a sentence.
- Using the TeX itself as a find target.
- Using tracked-change fragments as markers without accepting or inspecting the changes.

## Output contract

A successful render response reports paths and counts. It does not replace validation. A valid output must return `ok: true` from validation with:

- MathType object count at least equal to manifest equation count.
- Native number count equal to numbered equation count.
- Native reference count equal to manifest reference count.
- Sequential equation values beginning at 1.
- No unresolved markers or placeholders.
- No Word OMath objects under strict MathType output.
