# Contributing to Excel Formula Atlas

Thank you for helping make unfamiliar workbooks easier to understand. Useful reports, clear documentation, and small reproducible workbooks are as valuable as code changes.

## Report a problem

Please use the bug report template and include:

- Your Windows, Excel, and browser versions.
- Which version or commit you used, and whether you ran the distributed TXT or a development build.
- The action you took, the result you expected, and what actually happened.
- The exact error text, if any. For loading problems, include the stage and counts shown on the panel.
- A small fictional workbook that reproduces the issue, when possible.

Do not attach a work workbook or exported HTML containing private information. The HTML includes workbook values, formulas, labels, and metadata. Removing the original workbook from a report does not remove that data from its exported map.

## Propose an improvement

Describe the workbook task the improvement would make easier. For example, “I need to trace the inputs to several adjacent totals” is more useful than “add a new graph mode.” Include what you tried and which part of the current workflow gets in the way.

## Make a change

1. Keep the change focused on one behavior or problem.
2. Edit the source files and regenerate the distributed TXT using the repository's build tooling. Do not make a distribution-only fix that disappears on the next build.
3. Keep the viewer usable without external scripts, stylesheets, fonts, or network services. Live navigation may use the existing local connection; the standalone map must remain useful without it.
4. Keep English and Korean user-facing text consistent. Describe the user's action, rather than internal implementation details.
5. Record a user-visible change under **Unreleased** in [CHANGELOG.md](CHANGELOG.md).
6. Include validation evidence in the change description, along with any part you could not verify.

From the repository root:

```text
python scripts/build.py
python -m unittest discover -s tests -v
```

The package checks use Python's standard library and do not start Excel or a navigation listener. Rebuild the introduction page with `python scripts/build_site.py` after updating its template, embedded demo, or screenshots. This is development tooling; end users only need the distributed TXT.

### Optional example rebuilding

The included workbook and preview are ready to use. Rebuilding the example requires Node.js and an environment that can resolve `@oai/artifact-tool`. These are authoring dependencies, not requirements for macro users or the standalone HTML viewer.

```text
node scripts/build_example.mjs --qa-dir <outside-repo>
```

Replace `<outside-repo>` with a QA output directory outside this repository. The builder recalculates and checks all 197 formula results, tests changing `Assumptions!F8` from an 80% to a 90% collection rate, verifies the expected collections and closing-cash change with operating profit unchanged, restores the baseline input, and writes renders of all eight worksheets. Keep the calculation report and render files outside the public package.

The demo snapshot is translated preserved sample data cross-checked against the English workbook; it is not a fresh export from the latest VBA code. Rebuilding the workbook does not regenerate that snapshot. Follow [examples/README.md](examples/README.md) to keep the workbook, snapshot, and preview consistent, and report artifact-engine checks separately from native Excel tests.

## Validate the behavior you changed

Use a fictional sample and preserve its original copy. A check is useful when it can detect the specific regression your change might introduce.

### Viewer changes

- Open a newly generated map in a real browser and wait for it to finish loading.
- Check cell inspection, left-drag area selection, right-drag panning, wheel zoom, connection selection, search, and the nearby-table panel when those interactions might be affected.
- Check both languages, visible keyboard focus, and the layout at a narrow viewport.
- For rendering changes, test both many worksheets and one worksheet with many captured cells. Verify that the loading panel paints before heavy work, shows real counts, reaches completion, and reports failures.
- Confirm that static viewing works without starting the Excel navigation helper.

### VBA and reference-analysis changes

- Install the distributed TXT in a newly created standard module with no existing user code: choose **Insert → Module**, click that new module's code window, press **Ctrl + A**, and paste the entire TXT. This replaces an automatically inserted `Option Explicit` instead of duplicating it. Do not overwrite an existing module containing other macros.
- Compile and run the exported module in desktop Excel on Windows.
- Check current-workbook selection, file-picker selection, cancellation, output creation, and a failed analysis.
- Confirm that the source workbook is not edited and that Excel's application settings and status bar are restored after both success and failure.
- For parsing changes, compare the map with the actual formulas. Include a case that the previous code got wrong and a nearby case that should remain unchanged.
- Confirm that regenerated viewer and helper payloads match the intended source files.

### Navigation changes

- Check the helper's request handling and cleanup independently.
- Then perform the complete path: run the macro, open its connected page, click a cell, and verify that the correct workbook, worksheet, and cell are selected in Excel.
- Report a component check as a component check. Do not describe it as a verified browser-to-Excel workflow.

Organizational policies may block macros or local connections. Record the actual error and the environment; do not recommend bypassing an organization's security policy to make a test pass.

## Review checklist

- Does the change help a person understand or navigate a workbook?
- Are exported snapshots and parsing limits described honestly?
- Can the distributed TXT still be pasted into one standard VBA module?
- Are real workbook data, private paths, credentials, and machine-specific files absent from the contribution?
- Do documentation and screenshots describe the final behavior?
