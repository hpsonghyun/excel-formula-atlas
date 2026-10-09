# Publishing the prepared candidate

This repository is prepared locally. No remote, public repository, GitHub Pages site or GitHub Release is required to use the files locally.

## Candidate contents

- `dist/ExcelFormulaAtlas.txt`: the complete pasteable VBA module, including its viewer and navigation helper.
- `src/`: editable VBA, viewer and helper sources.
- `docs/index.html`: the English introduction page.
- `docs/demo.html`: the interactive synthetic workbook map.
- `examples/OperationsModel.xlsx`: the editable synthetic example.
- `docs/downloads/`: copies served with the introduction page. The build keeps the VBA download identical to `dist`.
- `README.md`, `LICENSE`, `SECURITY.md`, contribution guidance and issue templates.

Run `python scripts/build.py` and the documented checks before selecting a candidate. Inspect the actual source, rendered pages and generated downloads. Keep local QA output and real workbook maps outside the repository. The publication review must distinguish browser tests, Excel macro tests and the full browser-to-Excel connection.

## Final publication decisions

Before describing this candidate as fully verified, finish the pending Windows Excel checks in [Validation status](docs/VALIDATION.md) against the exact TXT hash. The current browser checks do not establish native macro execution or the full Excel navigation connection.

1. Choose the GitHub owner and repository name. Review the exact commit and file inventory, including image pixels and workbook contents.
2. Confirm the README support wording and existing maintainer support URL, `https://buymeacoffee.com/namsonghyun2`. This project does not create a payment account or require a donation.
3. Decide whether to publish the source, a Release, Pages, or all three. These are separate actions.
4. For Pages, use the `main` branch and `/docs` directory, or an explicitly reviewed deployment workflow. The supplied introduction and downloads stay inside that publishing directory. Do not enable a deployment workflow merely to preview the files.
5. After authorized publication, check anonymous access to the repository, the exact downloadable TXT bytes, the live demo, images, links and mobile layout. Match Release assets to the reviewed local hashes.

GitHub recommends relative links for repository files, so the README works in a clone as well as on GitHub. [Relative links in GitHub Markdown](https://docs.github.com/en/get-started/writing-on-github/getting-started-with-writing-and-formatting-on-github/basic-writing-and-formatting-syntax#relative-links).

GitHub Pages accepts an `index.html` entry point in the selected publishing source. [Creating a GitHub Pages site](https://docs.github.com/en/pages/getting-started-with-github-pages/creating-a-github-pages-site).
