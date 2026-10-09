# Publishing Excel Formula Atlas

This document describes publication and verification of **0.1.1**. The release locations are the [source repository](https://github.com/hpsonghyun/excel-formula-atlas) and [introduction page](https://hpsonghyun.github.io/excel-formula-atlas/). Publication and anonymous access are checked separately from the dated artifact tests in the validation record. The downloaded files also work locally.

## Release contents

- `dist/ExcelFormulaAtlas.txt`: the complete pasteable VBA module, including its viewer and navigation helper.
- `src/`: editable VBA, viewer and helper sources.
- `docs/index.html`: the English introduction page.
- `docs/demo.html`: the interactive synthetic workbook map.
- `examples/OperationsModel.xlsx`: the editable synthetic example.
- `docs/downloads/`: copies served with the introduction page. The build keeps the VBA download identical to `dist`.
- `README.md`, `LICENSE`, `SECURITY.md`, contribution guidance and issue templates.

For each release, run `python scripts/build.py` and the documented checks. Inspect the actual source, rendered pages and generated downloads. Keep local QA output and real workbook maps outside the repository. Record browser tests, Excel macro tests and the full browser-to-Excel connection separately.

## Publication checks

The [Validation status](docs/VALIDATION.md) records the exact 0.1.1 TXT hash, completed checks, the user's report of working core functionality, and remaining native Windows verification limits. Keep those distinctions visible in release notes; the browser checks do not establish full-project VBA compilation or the full Excel navigation connection.

1. Review the exact commit and file inventory for `hpsonghyun/excel-formula-atlas`, including image pixels and workbook contents.
2. Retain the optional maintainer support URL, `https://buymeacoffee.com/namsonghyun2`. Using this project does not require a donation.
3. Record source publication, GitHub Release assets and Pages deployment separately.
4. For Pages, use the `main` branch and `/docs` directory, or a deployment workflow. The supplied introduction and downloads stay inside that publishing directory.
5. After publication, check anonymous access to the repository, the exact downloadable TXT bytes, the live demo, images, links and mobile layout. Match Release assets to the reviewed local hashes and record the verified publication date.

GitHub recommends relative links for repository files, so the README works in a clone as well as on GitHub. [Relative links in GitHub Markdown](https://docs.github.com/en/get-started/writing-on-github/getting-started-with-writing-and-formatting-on-github/basic-writing-and-formatting-syntax#relative-links).

GitHub Pages accepts an `index.html` entry point in the selected publishing source. [Creating a GitHub Pages site](https://docs.github.com/en/pages/getting-started-with-github-pages/creating-a-github-pages-site).
