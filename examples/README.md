# The fictional operations example

[OperationsModel.xlsx](OperationsModel.xlsx) is an editable, synthetic workbook for exploring Excel Formula Atlas. Its eight worksheets contain **197 formulas** linking orders, sales, costs, production, material purchases, and cash. The model is cross-linked but acyclic.

Every company name, customer, product, quantity, price, rate, and financial amount is fictional. The model uses an October 2026 scenario with monetary amounts in KRW. It is a teaching example, not financial guidance or a representation of a real business.

## Start at Summary!P30

**`Summary!P30` calculates closing cash:**

```excel
=SUM(P23:P25)-SUM(P26:P29)
```

It adds opening cash, current sales collections, and prior receivables collected, then subtracts material payments, production labor, fixed costs, and capital expenditure. It is a useful starting point because several worksheet branches meet in one result.

Select it in [the sample map](../docs/demo.html), follow the upstream sources, and compare the formula with the actual workbook. Profit and cash differ because collections, payments, production, and inventory follow different timing and quantity rules.

| Worksheet | What it contributes |
| --- | --- |
| **Summary** | Product results, operating profit, and closing cash. |
| **Assumptions** | Collection and payment rates, opening cash, fixed costs, and opening material stock. |
| **Sales** | Revenue by product, collections, and receivables. |
| **Costing** | Material and labor unit costs, then cost of goods sold. |
| **Purchasing** | Material requirements, purchases after opening stock, and current cash payments. |
| **Inventory** | Required production, closing inventory, and production labor paid. |
| **Orders** | Sixteen fictional orders, discounts, returns, net units, and revenue. |
| **Catalog** | Four products, three materials, prices, usage rates, and inventory targets. |

## What the files contain

- **`OperationsModel.xlsx`** is the plain Excel workbook. It has no embedded VBA project, external workbook links, or embedded objects.
- **`model-spec.json`** records the model's cells, formulas, expected results, worksheet regions, demo stops, and sensitivity-test expectations.
- **`map-data.json`** supplies the captured worksheet context and reference graph used by the browser demo.
- **`translations.json`** records the English translations used to prepare the preserved sample snapshot for this repository.

The public demo is assembled from `map-data.json` and the current HTML viewer template. Its snapshot is a **translated, preserved synthetic workbook snapshot**, cross-checked against the included English workbook. The checks compare all 197 captured formulas with the workbook and compare the workbook's cached formula results with the model's expected results; they also check graph endpoints and the closing-cash formula.

This provenance matters: **the supplied preview is not a new export produced by running the latest VBA analyzer in Excel.** Snapshot consistency, artifact-engine recalculation, browser interaction, native Excel macro execution, and complete browser-to-Excel navigation are different checks. See [Validation status](../docs/VALIDATION.md) for the current evidence.

## Optional rebuilding

You do not need to rebuild anything to use the example or macro. The included XLSX opens in Excel, and the generated HTML can be viewed directly.

To rebuild the example, a developer needs Node.js and an authoring environment in which the Node package **`@oai/artifact-tool`** can be resolved by `scripts/build_example.mjs`. The repository does not bundle that package. This optional authoring dependency is separate from the standard-library Python packaging scripts and is not required on the PC where a user pastes the VBA TXT.

From the repository root, run:

```text
node scripts/build_example.mjs --qa-dir <outside-repo>
```

Replace `<outside-repo>` with an actual folder outside the repository and quote the path if it contains spaces. If your environment cannot resolve `@oai/artifact-tool`, use the committed example and preview; the optional rebuild cannot run in that environment until the package is available.

The builder:

1. Creates the eight worksheets, formulas, and formatting, then recalculates them with `@oai/artifact-tool`.
2. Checks every one of the 197 formula results against an independently recorded expected value.
3. Changes **`Assumptions!F8` from 80% to 90%**. It checks that `Sales!Q27` becomes 90% of current revenue, that `Summary!P30` increases by 10% of current revenue, and that `Summary!K12` operating profit remains unchanged.
4. Restores the collection rate to 80%, recalculates, and verifies the baseline closing-cash result before exporting the workbook.
5. Writes `examples/OperationsModel.xlsx`, `examples/model-spec.json`, and the identical workbook download under `docs/downloads/`.
6. Writes calculation reports, a formula-error scan, and a PNG render of each worksheet to the QA directory for inspection.

Calculation and restoration failures stop the build. The calculation report identifies its engine and explicitly records that native Excel recalculation has not been verified by this builder. Inspect the renders as well as the numerical checks; passing calculations do not prove that labels or formatting are readable.

## Keep the example and preview consistent

The example builder does **not** analyze the XLSX with VBA or automatically replace `map-data.json`. If you change the workbook's formulas, labels, values, or worksheet geometry, review and update the snapshot too. For a fresh analyzer export, run the distributed macro against the synthetic workbook in desktop Excel and retain the resulting provenance when preparing a replacement public snapshot.

After updating and checking the snapshot, regenerate the demo and introduction page from the repository root:

```text
python scripts/build_demo.py
python scripts/build_site.py
python -m unittest discover -s tests -v
```

The Python checks read the XLSX's stored formulas and cached values. They do not launch Excel or recalculate the workbook. Review the regenerated map in a browser and compare key cells, ranges, and worksheet positions with Excel before replacing public screenshots or publishing a changed sample.
