# Validation status

This is a local, unpublished **0.1.0 release candidate**. Browser and artifact checks passed. Native Excel macro execution and the complete browser-to-Excel navigation path still require a Windows Excel verification run before a fully verified release.

## Exact candidate

Checked on 2026-10-09. SHA-256 identifies the bytes inspected:

| Artifact | SHA-256 |
| --- | --- |
| `dist/ExcelFormulaAtlas.txt` | `993760e77b62626b6671e7a5840f4b0901427b284783f4baa9825720adf85601` |
| `src/viewer/template.html` | `610c48739230792f65f494a80d9e2d03e02f4c28a43f90e3a4eaf9bd8d536745` |
| `examples/OperationsModel.xlsx` | `74f390d71714e76de8d5a0a3807df78cde68e8b223a12aa2ccc73a12c37e337a` |

Website downloads are byte-identical to their source artifacts. The TXT includes the complete MIT notice, analyzer, viewer, and navigation helper. It is one standard VBA module; no companion source files are needed by the macro user.

## Checks completed

The repository's Python test suite passed **35 checks**: 21 package checks, 7 synthetic-workbook checks, and 7 site/demo consistency checks. Browser interaction runs were additional checks.

| Scope | Evidence and outcome |
| --- | --- |
| Distribution | 21 package contract tests passed: one module, balanced procedures and strings, VBA line limits, English defaults, embedded payload equality, bounded local navigation, read-only source opening, real progress counters, and MIT notices. These are static checks; they do not compile VBA. |
| Rebuild | A build from an isolated copy of the repository inputs reproduced the exact TXT bytes. JavaScript syntax checks passed for both viewer scripts. |
| Synthetic workbook | All 197 formulas were recalculated with the artifact engine and compared with independently computed expectations. The saved XLSX formulas and cached values were also checked. Collection rate 80% to 90% changed closing cash from 48,267,840 to 51,389,420 while operating profit stayed unchanged; the input and results were restored before saving. |
| Workbook integrity | Seven demo contract tests passed: workbook/download equality, plain XLSX with no VBA or external links, English synthetic data, formula/cache agreement, snapshot formulas, graph endpoint integrity, and the closing-cash result. All eight worksheet renders were visually reviewed. |
| Browser interactions | Native mouse actions passed search, single-area selection, larger-area selection, right-button pan, and line selection. Detail-cell clicks, upstream depth, and English/Korean changes passed. Selecting the cash area produced 10 seeds, 21 upstream items, and 5 related sheets. Right-button pan preserved the selection. |
| Loading | The 8-sheet example, an 80-sheet stress fixture, and a one-sheet 20,000-cell fixture completed at 100%. The loading panel was visible on the first observed animation frame, counters stayed within their totals, percentages were monotonic, drawing yielded between chunks, and the panel closed when ready. Invalid embedded JSON kept the error panel visible with an explanation. |
| Introduction | The final page rendered in a browser, its embedded map opened and selected `Summary!P30`, all four screenshot placements decoded, and macro-name copy worked. Desktop and 390/320-pixel layouts were checked for horizontal overflow. No HTTP requests or browser errors occurred while loading and using the preview. |

Browser checks loaded the exact self-authored HTML into an isolated browser page. They did not start a local HTTP server. The sample map has no live Excel connection. Its snapshot is a translated, preserved synthetic fixture cross-checked against the included English workbook, rather than a fresh export from the latest VBA.

Loading percentages measure completed preparation items. Initial parsing remains indeterminate and may briefly block the browser; the stress runs included a parsing task of about 0.2 seconds. Progress is not a prediction of time remaining or a guarantee for every workbook size.

## Verification still pending

- Paste the exact current TXT into desktop Excel, compile the module, and execute analysis using both source choices.
- Confirm the current workbook's formulas, values, hidden-sheet state, and Excel application settings survive analysis unchanged.
- Run the generated HTML with the local helper in an allowed Windows environment and confirm a cell click selects the exact workbook, worksheet, and range in Excel.
- Confirm stopping the connection, source closure, and unavailable-connection messages in a real session.

The available Windows Excel COM instance could not be attached to, and a single new-instance creation attempt did not return within its bounded diagnostic interval. Existing user workbooks were not accessed or closed. No Trust Center, registry, execution-policy, or organizational policy settings were changed. This environment failure is not proof that the macro succeeds or fails in Excel. Results from earlier builds are not used as proof for this TXT.

## Finish the Windows check

For a manual check, use a disposable macro-hosting workbook and the fictional sample:

1. Create a new standard module containing no user code in the Visual Basic editor. Click inside that new module's code window, press **Ctrl+A**, and replace its contents with the entire `dist/ExcelFormulaAtlas.txt`. This replaces any automatically inserted `Option Explicit`. Preserve existing personal macro modules. Save the disposable host as `.xlsm`, select its project in Project Explorer, then use **Debug → Compile VBAProject** and confirm no compile error appears.
2. Open `examples/OperationsModel.xlsx`, make it active, and run `ExcelMap_Start`. Choose **1 — Current workbook** and an output folder. Confirm eight sheets, 197 formulas, and `Summary!P30 = 48,267,840` in the exported map.
3. Repeat with **2 — Choose an Excel file** using a separate copy of the fictional example. Confirm the newly opened analysis source is read-only. Compare its formulas and values before and after analysis.
4. If the local helper starts under your existing policy, use **Open this location in Excel** and the automatic navigation option on several exported cells on different sheets. Confirm Excel's workbook name, worksheet tab, and name box agree with the map. A context-only cell is intentionally unavailable for this action.
5. Test right drag, left-area selection, line selection, language changes, and loading in that generated HTML. Run `ExcelMap_Stop` and confirm the static map remains usable while live navigation reports disconnection.
6. Record Excel version, Windows version, TXT hash, successful checks, and any error messages. Do not include real workbook data in a public issue.

An optional automation harness, `tests/windows_excel_smoke.ps1`, is prepared to check module insertion and readback, both analysis sources, HTML generation, application-state restoration, and the navigation queue-to-Excel component. Its PowerShell syntax was checked; it has **not** been executed for this candidate. Running selected procedures does not establish full-project compilation, so the manual Compile step remains separate. The harness requires the existing environment to permit VBA project automation and never enables that permission itself. It does not call the HTTP helper startup entry points or prove the complete browser connection.

Only mark the pending items verified after a real run against the exact candidate, and update this document with the new evidence before publication.

Microsoft documents automatic `Option Explicit` insertion under [Require Variable Declaration](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/options-dialog-box) and full-project compilation under the [Debug menu](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/debug-menu).
