# PDF_BULK_FILL

Excel-driven bulk PDF filler supporting:
- **Native PDF AcroForm Comb Fields & Checkboxes** (one letter per box, e.g., `RR-Form.pdf`)
- **Employee ID Card Image Overlays** (front & back, auto-scaled and centered)
- **Coordinate-Based Text Overlays** (full backward compatibility with legacy `TEMPLATE.xlsx`, encrypted `.xlsx` workbooks, `!<CONCAT>`, and `SrinkToFit`)

## Quick Start

```bash
# Install in a virtual environment
python3 -m venv .venv
source .venv/bin/activate
pip install -e ".[dev]"

# Run headless CLI with RR-Form.pdf
pdf-bulk-fill \
  --pdf templates/pdf/RR-Form.pdf \
  --template templates/excel/RR-Form_TEMPLATE.xlsx \
  --source-dir data/input \
  --image-dir data/input/ids \
  --output-dir data/output \
  --flatten

# Or run using a saved session.json
pdf-bulk-fill --session session.json

# Or launch interactive Tkinter file/password dialogs
pdf-bulk-fill --gui
```

## Running Tests & Coverage
```bash
pytest -v
```
All commits require passing tests with **≥ 80% code coverage** (enforced automatically by `pyproject.toml`).

## Documentation
- [Project Rules](rules.md)
- [Requirements](docs/REQUIREMENTS.md)
- [Project Board](https://github.com/users/vipulasrilanka/projects/2)
