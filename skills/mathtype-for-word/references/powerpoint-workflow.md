# PowerPoint workflow

## Supported result

Create editable `Equation.DSMT4` OLE equations in `.pptx` without showing Word, PowerPoint, or MathType. Hidden Word converts one-line TeX through MathType's `MTCommand_TeXToggle`; the bridge copies the resulting OLE object into a windowless PowerPoint presentation and verifies its MathML against the manifest. MathType 7 inserts external objects as floating shapes because PowerPoint does not support inline external objects.

MathType's Word-only `MTPlaceRef`, `SEQ MTEqn`, `MTReference`, and `GOTOBUTTON/REF` workflow does not exist in PowerPoint. Do not emulate it with typed equation numbers or fake hyperlinks.

This is silent COM/OLE automation. Word, PowerPoint, and MathType remain hidden, and the bridge does not use mouse, keyboard, focus, `AppActivate`, or `SendKeys`. It briefly uses the Windows clipboard to transfer the OLE object, then validates the returned MathML before saving the PPTX.

## Presentation manifest schema v1

```json
{
  "schema_version": 1,
  "equations": [
    {
      "id": "energy_relation",
      "marker": "{{MATH:energy_relation}}",
      "tex": "E=m c^2",
      "mathml": "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><mi>E</mi><mo>=</mo><mi>m</mi><msup><mi>c</mi><mn>2</mn></msup></mrow></math>",
      "height_points": 32
    }
  ]
}
```

Rules:

- Use unique stable ASCII IDs.
- Supply one-line MathType-compatible `tex` for silent creation and one complete Presentation MathML `<math>` element for independent validation. Verify that both encode the same operators, scripts, fractions, radicals, accents, and identifiers.
- Put each marker alone in one top-level text box. Grouped markers and markers embedded in prose are rejected.
- Equation size follows the Word rule — one uniform math size, with the height following the content:
  - Omit `height_points`. The math size is taken from, in order, the equation's `font_pt`, the manifest-level `equation_font_pt`, the font size of the marker text box, else 24 pt. The natural MathType object (math at the preference file's Full size: 10.5 pt for CJGE, 12 pt with equation_preferences "none") is scaled by `font_pt / Full`.
  - Each object records `MT_SIZE_MODE`, `MT_NATIVE_HEIGHT_PT` and `MT_FONT_PT` tags; `validate_mathtype_powerpoint_presentation` fails if an equation was resized away from its math size and warns when a deck mixes sizes.
  - `height_points` (12–200) remains as a legacy fixed-height override. Avoid it: it gives every equation the same height and therefore different glyph sizes.
- Write each marker in the same font size as the surrounding prose so equations match the text.
- If PowerPoint is already open, the bridge attaches to it (PowerPoint is single-instance) and closes only its own presentation; it never quits the user's PowerPoint.
- Keep descriptive prose in separate text boxes above and below the marker.

## Equation preferences in PowerPoint

`render_mathtype_powerpoint_presentation` re-typesets each converted equation with the same MathType preference file as Word (manifest `equation_preferences`, default CJGE; `"none"` to skip) before it is placed on the slide. The natural object then corresponds to the preference file's Full size (10.5 pt for CJGE), recorded in the `MT_NATIVE_FONT_PT` tag, and the object is scaled to the deck's math size (`font_pt`).

## Direct bridge commands

```powershell
# Resolve these variables from the user's attached or repository files; never reuse a sample path literally.
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File scripts\mathtype-word.ps1 `
  -Action probe-pptx

pwsh.exe -NoProfile -ExecutionPolicy Bypass -File scripts\mathtype-word.ps1 `
  -Action render-pptx `
  -InputPath $inputPptx `
  -OutputPath $outputPptx `
  -ManifestPath $manifestJson

pwsh.exe -NoProfile -ExecutionPolicy Bypass -File scripts\mathtype-word.ps1 `
  -Action validate-pptx `
  -InputPath $outputPptx `
  -ManifestPath $manifestJson
```

## Validation gate

Require all of the following before claiming completion:

- One named `Equation.DSMT4` OLE object for every manifest equation.
- Object name `MathType_<id>` and alternative text identifying the equation.
- Horizontal centering within one point.
- No unresolved `{{MATH:...}}` marker.
- Full-size rendered slide inspection with no clipping, overlap, or illegible equation sizing.
- No Word, PowerPoint, or MathType process created by the render action remains after completion.
