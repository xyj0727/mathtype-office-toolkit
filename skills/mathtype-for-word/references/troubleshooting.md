# Troubleshooting

| Symptom | Likely cause | Action |
|---|---|---|
| `word_com` is false | Word desktop missing, broken registration, or blocked automation | Repair/install Word; retry the probe in the interactive user session. |
| MathType executable or template missing | MathType or Word add-in is not installed | Repair MathType and enable its Word support. |
| `equation_dsmt4_registered` is false | MathType OLE registration is broken | Run MathType repair as the same Windows user. |
| Toggle TeX creates no object | Unsupported delimiter/TeX, add-in not loaded, or interactive dialog | Use one-line portable TeX; confirm `$...$` for inline and `\[...\]` for display; inspect Word add-ins. |
| Toggle TeX creates Word OMath | Word conversion was used instead of MathType | Remove the OMath result and rerun through `MTCommand_TeXToggle`. |
| Number is `(1.1)` | Existing MathType format includes section number | Render through the bridge; it removes current `MTSec`/`MTChap` components while preserving `MTPlaceRef`. |
| `MTReference` remains | The target `MTPlaceRef` action did not complete | Do not type over it; retry against a native numbered equation. |
| `Error! Reference source not found.` | A `ZEqnNum...` bookmark was deleted or the field was copied incorrectly | Recreate the reference through MathType; do not hand-edit the bookmark. |
| Output is locked | Word or another process has the DOCX open | Close it or choose a new output path. |
| Word remains in Task Manager after a failure | A COM call or MathType dialog blocked | Identify only the Word process started by the failed job before terminating it; never kill all Word sessions. |
| MCP returns a bridge timeout | A Word/MathType COM macro blocked longer than the watchdog | Check `isolated_word_pid` and `isolated_word_process_terminated`; confirm preferences were restored, then retry one small fixture. The default watchdog is 240 seconds and may be changed with `MATHTYPE_WORD_TIMEOUT_SECONDS`. |
| Validation reports OMath | Existing or newly inserted built-in Word math is present | Convert the requested equations to MathType; inspect unrelated legacy OMath separately before deletion. |
| Render hangs at `MTCommand_InsertEqnNum`; hidden Word shows an "Insert Equation Number" dialog | `HKCU\Software\Design Science\DSMT7\WordCommands` values are REG_DWORD (or the key is missing, typically after reinstalling Office). The MathType Word add-in reads `NoEqnNumWarningDlg` / `NoInsertEqnRefDlg` as REG_SZ strings | Run `configure_mathtype_word_defaults` (it now writes REG_SZ `"0"`/`"1"`), or recreate both values as strings. Never change them while a render is running. |
| Render fails with "property `Content` not found" right after opening the DOCX | The document is under `%TEMP%`, so Word opens it in Protected View | Move the input and output into a normal folder (for example Documents). |
| Bridge error text is garbled / "invalid JSON" | Localized (e.g. Chinese) error text written in the ANSI code page | Fixed: the bridge now writes UTF-8. |
| The user's PowerPoint closed after a PPTX job | PowerPoint is single-instance and the bridge attached to it | Fixed: the bridge now closes only its own presentation when PowerPoint was already running. |
| MathType's *Format Equations* fails in Word with error 53 `MathPage.WLL` after reinstalling Office | `MathPage.wll` is missing from Word's startup folder | Copy `<MathType>\MathPage\64\MathPage.wll` (64-bit Office) into `%APPDATA%\Microsoft\Word\STARTUP` and restart Word. The toolkit itself calls the MathType API directly and does not need it in Word. |
| `equation_format` errors in validation | Equations are not typeset with the CJGE preferences (edited or added later) | Run `apply_mathtype_equation_preferences`, then validate again. |
| MCP server emits parse errors | A launcher/log wrote to stdout | Protocol output must be JSON only; keep diagnostics on stderr. |

## Recovery rules

- Keep the input untouched.
- Record the equation ID, marker, TeX, and exact bridge error.
- Check for a blocking dialog in Word/MathType if a call times out.
- Clean only the isolated Word process created by the failed job.
- Confirm temporary warning registry values were restored.
- Rerun the smallest fixture before rerunning a large document.

## Fallback boundary

Word captions, numbered lists, manually typed numbers, and Word OMath are a fallback only. Before using one:

1. Show the failed prerequisite or conversion result.
2. Explain that the output will not be MathType-native.
3. Obtain explicit user acceptance.
4. Label the deliverable as a fallback and do not claim native MathType references.
