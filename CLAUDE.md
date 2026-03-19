# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Python application to automate creation of purchase orders (Ordens de Compra) for the M8 ERP system. Integrates with the `proz-m8` REST API client library.

## Setup

```bash
# Activate virtual environment
source venv/bin/activate

# Run the application
python src/main.py
```

## Dependencies

- `openpyxl` — Excel file reading/writing
- `proz-m8` (installed from git) — M8 API client library (`m8` package)

## Language

All user-facing strings and messages must be written in **Brazilian Portuguese (pt-BR)**. Everything else (code, logs, and comments) must be in **English**.


## Architecture

**`src/main.py`** — Entry point (currently placeholder).

**`src/utils/excel_helper.py`** — `ExcelHelper` class wrapping `openpyxl` for reading/writing Excel files: cell operations, formatting, row/column manipulation, formulas, merging.

**`src/utils/oc_loader.py`** — Loads purchase order data from the "OC" sheet in an Excel file, returning a list of items.

**`m8` package (in venv)** — External library providing:
- `M8` class — REST API facade with `@auth` decorator for automatic token management. Methods for purchase orders, invoices, and receivables.
- `PurchaseOrder`, `PurchaseOrderItem`, `PurchaseOrderInstallment` — Dataclasses with `to_dict()` for API serialization.
- `load_credentials()` — Loads API credentials from file.

## Typical Data Flow

1. Read purchase order data from an Excel file via `oc_loader.py`
2. Instantiate `PurchaseOrder` / `PurchaseOrderItem` dataclasses
3. Use `M8` client to POST the order to the ERP API
