# Thinkswiss

## Cursor Cloud specific instructions

### What this repo is
This is **not** a long-running application/service. It is a collection of one-off
**Python document-generation scripts** (Nepal policy / environmental content). The
`main` branch only contains this README/AGENTS file; the actual scripts live on
feature branches (e.g. `build_excel*.py`, `make_presentation.py`,
`translate_*.py`, `curriculum_builder.py`). Each script builds an output document:

- `.xlsx` spreadsheets via **openpyxl**
- `.pptx` presentations via **python-pptx** (imported as `pptx`)
- `.docx` documents via **python-docx** (imported as `docx`)

### Running a script ("the application")
There is no build/serve step. Run a script directly with Python 3:

```
python3 <build_script>.py
```

Dependencies (`openpyxl`, `python-pptx`, `python-docx`) are installed by the
startup update script, so no per-run install is needed.

### Gotchas
- **Scripts hardcode `/workspace/...` output paths.** Running a script writes its
  output document into the repo root (`/workspace`), producing an untracked file.
  Move or delete the generated artifact afterward if you want a clean working tree.
- There are **no automated tests, linters, or CI config** in this repo. "Verifying"
  a script means running it and (optionally) reading the output back with the same
  library (`openpyxl.load_workbook`, `pptx.Presentation`, `docx.Document`).
- `pip3 install` resolves to a **user install** (`~/.local`); no venv/PEP-668 issues.
