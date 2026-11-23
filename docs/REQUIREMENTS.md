# Requirements — PDF_BULK_FILL

## 1. Objective
Build a clean, extensible Python application to bulk-fill PDF documents from Excel data sources, supporting:
1. **Native PDF AcroForm Comb Fields & Checkboxes** — e.g., `RR-Form.pdf`, where each fillable box holds exactly **one character**, plus checkbox selection (`Internet`, `SMS`).
2. **Image Placement** — inserting two employee ID card images (**front** and **back**) cleanly into the top area of `RR-Form.pdf`.
3. **Legacy Coordinate-Based Text Overlays** — full feature parity with the existing overlay workflow (`TEMPLATE.xlsx`, multi-workbook lookups, encrypted Excel files, preprocessors, `!<CONCAT>`, `SrinkToFit`, `session.json`, and Tkinter pickers).

---

## 2. Functional Requirements

### FR-1: Excel Data & Encrypted Workbook Support
- Read record lists (`Data` sheet) and mapping definitions (`Overlay` sheet) from a `TEMPLATE.xlsx` workbook.
- Support looking up values across multiple external `.xlsx` source files by primary key (supporting both numeric and alphanumeric keys).
- Detect and decrypt password-protected `.xlsx` workbooks (`msoffcrypto`) using temporary files that are always cleaned up after execution.
- Index loaded sheets once in memory for fast batch lookups.

### FR-2: Template Parsing & Text Preprocessing
- Preserve compatibility with `<key=value>` template definitions, immediate text values, and `!<CONCAT><TargetName>` concatenation.
- Support all existing preprocessors:
  - `changeTextCase` (`UPPER`, `lower`, `Title`)
  - `FormatDate`
  - `NumberToText` (`Integer`, `Floating Point`)
  - `NumberToCurrency` (with Unicode CLDR/Babel currency code & symbol recognition on either side, `Decimal` precision, locale-aware decimal/grouping separators, currency mismatch & ambiguity rejection, optional template-level `Settings` sheet defaults and per-field overrides)
  - `FormatNumber` (same `Decimal` & locale-aware numeric validation)
  - `AddSpace`
- Support fixed-box comb line splitting (e.g., splitting **Full Name** across `22 / 31 / 31` boxes and **Name with Initials** across `20 / 31` boxes without breaking words awkwardly).

### FR-3: AcroForm Comb Fields & Checkboxes (`RR-Form.pdf`)
- Populate native AcroForm fields by field name, respecting each field's `MaxLen` and comb spacing (1 character per box):
  - `Serial No` (text)
  - `Zone` (1), `Emp No` (6), `Mem No` (6), `NIC` (12)
  - `Name 1` (22), `Name 2` (31), `Name 3` (31)
  - `Name INI1` (20), `Name INI 2` (31)
  - `Mobile` (10), `Email` (27)
  - `Internet` (checkbox), `SMS` (checkbox)
- Render crisp appearances so every character sits centered inside its box across macOS Preview, Adobe Acrobat, web browsers, and printing.
- Support both editable and flattened (print-ready) PDF output modes.

### FR-4: Generic DocType-Matched Data-Driven & Constant Image Overlays
- Support two line-by-line `Type=Image` modes in `TEMPLATE.xlsx` (`Overlay` sheet):
  1. **Data-Driven Images** (e.g., Old & New Sri Lankan NIC front/back):
     - `TEMPLATE.xlsx` defines slot geometry (`X`, `Y`, `Width`, `Height`, `Origin`) and accepted `DocType` in `Param`, plus the source workbook/column in `Content`.
     - The data workbook (`employees.xlsx`) stores image asset properties (`<DocType=...><FileName=...><Location=...><Rotation=...><KeepAspect=...>`, with defaults `Location=""`, `Rotation=0`, `KeepAspect=True`).
     - If `image.DocType != slot.DocType` (or cell is empty), log a warning and skip the row cleanly.
  2. **Constant / Static Images** (e.g., company logo or rubber stamp):
     - `TEMPLATE.xlsx` defines `FileName`, optional `Location`, `Rotation`, `KeepAspect` directly in `Content` and placement (`X`, `Y`, `Width`, `Height`, `Origin`) in `Param`.
- Apply EXIF orientation normalization and explicit clockwise `Rotation`, preserve aspect ratio when `KeepAspect=True`, and support common formats (`.jpg`, `.jpeg`, `.png`, `.bmp`, `.webp`).

### FR-5: Legacy Coordinate-Based Text Overlays
- Support drawing text at explicit `(X, Y)` coordinates with `pt`, `mm`, and `in` units.
- Automatically detect each PDF template's actual page dimensions (A4, Letter, Legal, etc.).
- Support `Font`, `FontSize`, `LineSpace`, and `SrinkToFit` (word-wrap + font-size reduction within max width/lines and optional first-line indent).

### FR-6: Batch Execution, Session Persistence & UI
- Support both interactive GUI mode (Tkinter file pickers & password prompt) and headless CLI mode (`session.json` / CLI arguments).
- Save and reload `session.json` for repeatable runs.
- Generate one output PDF per record named `<key>-<identifier>.pdf` in the configured output folder, reporting per-record success/failure clearly.

---

## 3. Non-Functional Requirements
- **Code Quality & Modularity**: Clean Python 3.10+ package structure with type hints and clear separation of concerns (Excel loader, template parser, preprocessors, PDF engine, CLI/GUI runner).
- **Testing & Coverage**: Automated `pytest` suite covering all modules with **≥ 80% code coverage** enforced before commit.
- **No External Side Effects**: Temporary decrypted files must always be deleted (even on errors).
