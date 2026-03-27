# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Python PyQt6 desktop application to automate creation of purchase orders (Ordens de Compra) for the M8 ERP system. Integrates with the `proz-m8` REST API client library.

## Setup and Running

```bash
# Activate virtual environment
source venv/bin/activate

# Run the application
python src/app.py

# Run tests
pytest

# Run a single test file
pytest tests/test_oc_loader.py

# Run a single test
pytest tests/test_oc_loader.py::test_validate_file
```

## Dependencies

- `openpyxl` — Excel file reading/writing
- `PyQt6` — GUI framework
- `proz-m8` (installed from git) — M8 API client library (`m8` package)

## Language

All user-facing strings and messages must be written in **Brazilian Portuguese (pt-BR)**. Everything else (code, logs, and comments) must be in **English**.

## Architecture

**`src/app.py`** — Entry point: creates `QApplication`, instantiates `MainWindow`, runs the event loop.

**`src/gui/main_window.py`** — GUI layer:
- `MainWindow` — Main window with import/create buttons, OC table, progress bar. Handles file selection, validation, table population, and status color-coding (green = success, red = error).
- `CriarOCWorker` — `QThread` subclass that calls the M8 API in the background, emitting signals for progress and per-item results.
- `CredenciaisDialog` — Login dialog for M8 credentials.
- `SobreDialog` — About dialog rendering `static/about.md` as markdown.

**`src/utils/oc_loader.py`** — Business logic for purchase order creation:
- `validate_file()` — Validates Excel structure; requires a "Cadastro" sheet with version string "v2".
- `load_oc()` — Reads the "OC" sheet, returns a list of dicts (skips empty rows).
- `generate_purchase_order()` — Maps spreadsheet fields (unit, company, type) to M8 IDs (company, employee, cost center, fiscal operation). Accepts `tipo_oc`: `"professor"` or `"preceptor"`.
- `generate_purchase_order_bulk()` — Drives bulk creation from the list returned by `load_oc()`.

**`src/utils/excel_helper.py`** — `ExcelHelper` class wrapping `openpyxl`: cell/range read-write, formatting, row/column insert/delete, merge, formulas, save/load.

**`m8` package (in venv)** — External library:
- `M8` — REST API facade with `@auth` decorator for automatic token management.
- `PurchaseOrder`, `PurchaseOrderItem`, `PurchaseOrderInstallment` — Dataclasses with `to_dict()` for API serialization.
- `load_credentials_from_file()` — Loads API credentials from file.

**`static/`** — `about.md` (displayed in the About dialog) and the template Excel file users download to fill OCs.

## Typical Data Flow

1. User imports an Excel file → `validate_file()` checks structure/version, `load_oc()` returns rows
2. `generate_purchase_order_bulk()` maps each row to `PurchaseOrder` dataclasses
3. Table is populated; user reviews and enters M8 credentials
4. `CriarOCWorker` iterates the list, calling `M8.create_purchase_order()` for each, reporting results back to the UI
