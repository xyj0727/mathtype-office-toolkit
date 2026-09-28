# CJGE MathType equation format (default)

The toolkit's default MathType format follows the *Chinese Journal of Geotechnical Engineering* (《岩土工程学报》, CJGE): measured from a published CJGE paper (2026, 48(9), DOI 10.11779/CJGE20250567, equations (1)–(18)) and the journal's author guidelines (https://www.cgejournal.com/zhenggaojianze). Every rule below is enforced or checked by the tools unless marked *writing rule* (the agent applies it while writing the manuscript and manifest).

## 1. Sizes and styles (MathType preferences)

Stored in `config/cjge_equation_preferences.eqp`, applied by `apply_mathtype_equation_preferences` (automatically inside `render_mathtype_word_document` and `render_mathtype_powerpoint_presentation`), and checked by `validate_mathtype_word_document`.

| MathType setting | Value |
|---|---|
| Full | **10.5 pt** (五号, same as CJGE body text) |
| Subscript / Superscript | **58 %** (≈ 6.1 pt) |
| Sub-subscript | 42 % |
| Symbol / Sub-symbol | 150 % / 100 % |
| Text | Times New Roman, upright |
| Function (sin, tan, max, ln, exp …) | Times New Roman, **upright** |
| Variable | Times New Roman, **italic** |
| Lower-case Greek | Symbol, **italic** |
| Upper-case Greek | Symbol, upright |
| Symbol (operators, brackets, braces) | Symbol, upright |
| Vector-Matrix | Times New Roman, **bold italic** |
| Number | Times New Roman, upright |

Long equations: always edit at 10.5 pt. If an equation does not fit the column (single column 8.53 cm), split it over lines or ask for a full-width layout; do not shrink the font.

## 2. Characters (*writing rule*, expressed in the TeX of the manifest)

| Element | Style | TeX |
|---|---|---|
| Variables | italic | `H`, `L`, `W` |
| Greek variables | italic | `\alpha`, `\beta`, `\theta`, `\varphi` |
| Descriptive subscripts (t, b, d, u, l, r, s, e, a, c, max …) | **upright** | `W_{\mathrm{t}}`, `\beta_{\mathrm{d}}`, `d_{\max}`, `C_{\mathrm{u}}` |
| Variable subscripts (index i, j) | italic | `x_{i}` |
| Numeric subscripts | upright | `L_{1}` |
| Function names | upright | `\tan`, `\max`, `f^{-1}(\cdot)` (the −1 is upright) |
| Units | upright | `\mathrm{m}`, `\mathrm{kPa}`, `\%` |
| Minus sign | Symbol minus “−” | write `-` in TeX; MathType stores U+2212 |
| Orders of magnitude | × and superscript | `2.0\times 10^{7}\ \mathrm{m}^{3}` |

## 3. Display layout

| Element | Format |
|---|---|
| Position | own line, centred in the column |
| Number | half-width `(1), (2) …`, Times New Roman 10.5 pt upright, right-aligned to the column edge; one sequence for the whole paper |
| Layout | 1×3 borderless table (`display_layout: "table"`, default; `config/cjge_layout_profile.json`): side cells 72 pt, zero padding, "at least" 15.6 pt lines, no space before/after |
| Punctuation | “，” when “式中：” follows, “。” when the sentence ends; placed between equation and number in 宋体 (`"punctuation": "，"` per equation) |
| Simultaneous equations | stacked rows, one right brace, **one number**: `\left.\begin{array}{l}W_{\mathrm{t}}=L\tan\beta_{\mathrm{d}}\\ V_{\mathrm{L}}=W_{\mathrm{b}}H\end{array}\right\}` (keep a space after `\\`; `aligned` and `\cr` are not supported by MathType's TeX import) |
| Line spacing | body text uses fixed 15.6 pt; equation paragraphs use "at least" so tall equations are not clipped |
| Sub-numbers (16a), (16b) | not generated automatically; number the group once or edit in MathType |

## 4. References in the text

- Written as **式（5）** or **见式（16）**: “式” plus **full-width** brackets (`reference_brackets: "fullwidth"`, default). The number inside is still a live MathType reference that follows renumbering.
- Write the “式” in the prose before the `{{EQREF:…}}` marker: `由式{{EQREF:r1}}可得` → 由式（1）可得.

## 5. Symbol explanations “式中：” (*writing rule*)

- Directly below the equation, new paragraph, **no indent**, starting with “式中：”, 宋体 10.5 pt.
- Items separated by the full-width semicolon “；”, the last item ends with “。”.
  Example: “式中：*R*<sub>a</sub>为试验值；*R*<sub>c</sub>为计算值；*R*<sub>e</sub>为相对误差。”
- Symbols already explained may be covered by “各参数意义同上”.
- A symbol in the prose must look exactly like the same symbol in the equation (font, italic/upright, subscripts).

## 6. Symbols in running text (*writing rule*)

- Simple symbols in prose are typed as text, not MathType objects: Times New Roman italic 10.5 pt, subscripts with Word's subscript format (≈ 7 pt), e.g. “滑距 *S*”, “坐标系 *xyz*”, *W*<sub>t</sub>.
- Use an inline MathType equation only for real expressions (fractions, radicals, sums …).
- `scan_plain_text_math` with `strategy: "cjge"` (default) proposes exactly this split and `prepare_mathtype_markers` writes the italic text and Word subscripts. If the user explicitly wants every symbol as a MathType object, use `strategy: "all"` and mention that this departs from the CJGE rule.
- Paragraphs holding inline MathType need "at least" instead of exact line spacing, or the top of the equation is clipped; render applies this automatically (`inline_line_spacing`).

## 7. Checklist

- [ ] Full 10.5 pt, sub/superscripts 58 % — `validate_mathtype_word_document` (equation_format).
- [ ] Variables TNR italic, functions upright, lower-case Greek Symbol italic, vectors/matrices bold italic — same check.
- [ ] Descriptive subscripts upright (`\mathrm{}`), index subscripts italic.
- [ ] Numbers `(n)` right-aligned, continuous.
- [ ] “，” or “。” after each display equation.
- [ ] “式中：” unindented, items separated by “；”, ending with “。”.
- [ ] Same symbol written identically in prose and equations.
- [ ] Minus is “−”, not “-”.
