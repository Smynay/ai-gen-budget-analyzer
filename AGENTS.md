# budget

Personal finance repository. Source PDF bank statements are in `sources/`.

## Source data

- `sources/Vypiska_po_schyotu_debetovoy_karty.pdf` — debit card account statement (Russian, 117p, Sberbank)
- `sources/Spravka_o_dvizhenii_sredstv.pdf` — funds movement statement (Russian, 30p, Tinkoff)

## Extraction

Run `.venv\Scripts\python.exe extract.py` to extract all PDFs in `sources/` into `results/YYYY-MM-DD_N/excel.xlsx`.

- **Config**: `config.yaml` — defines the unified schema, column templates, and matching rules. Categories can be overridden in `config_personal.yaml` (not tracked by git)
- **Virtual env**: `.venv\` (created by `uv venv`)
- **Output**: `results/YYYY-MM-DD_N/excel.xlsx` — Сводный sheet (all transactions + source column) + one sheet per PDF (without source)
- **Column order**: Дата, Сумма, Валюта, Категория, Описание, (Источник)
- **Amounts**: Sberbank unsigned = expense (negative), Tinkoff uses +/- signs
- **Time parsing**: both PDFs support `HH:MM` in transaction dates; stored as `DD.MM.YYYY HH:MM`
- **Continuation rows**: multi-row descriptions are merged and deduplicated; footer text ("Продолжение на следующей странице") is filtered out
- **Categorization**: keyword-based first pass via `config_personal.yaml` → `categories` (or `config.yaml` as fallback). Optional `direction` field (`income` / `expense`) limits matching by transaction sign. Unmatched → `Неизвестно` (fallback).

## AI Categorization

Run `.venv\Scripts\python.exe categorize.py --dump` to group unknown transactions from the latest XLSX into `results/YYYY-MM-DD_N/descriptions.yaml`. Each group contains similar descriptions (after normalizing numbers, known prefixes, etc.).

1. `categorize.py --dump` — creates `descriptions.yaml` with groups of similar unknown transactions
2. AI fills `category` for each group in `descriptions.yaml`
3. `categorize.py --apply` — reads the YAML and updates categories in the XLSX

Categories placed on individual group entries; if a group's `category` is set to `Неизвестно` or left `null`, those transactions remain uncategorized.

## Report

Run `.venv\Scripts\python.exe report.py` to generate an interactive HTML report from the latest XLSX.

Supported filters:
- `--year YYYY` — only transactions from that year
- `--month YYYY-MM` — only transactions from that month
- `--from YYYY-MM-DD [--to YYYY-MM-DD]` — arbitrary date range (inclusive)

Output: `results/YYYY-MM-DD_N/report.html` (or `report-YYYY.html`, `report-YYYY-MM.html`, `report-YYYY-MM-DD_YYYY-MM-DD.html`).

Features: Chart.js pie chart, Расходы/Доходы/Всё toggle, category table with drill-down, sortable date column, toggleable source column.

### Templates
Templates in `config.yaml` match PDFs by keywords in metadata/first-page text. Each defines column x-ranges and amount format (decimal/group separator). Supported:
- `sberbank_debit` — for Sberbank выписка по счёту
- `tinkoff_movement` — for Tinkoff справка о движении средств

To add new bank types: use the `add-bank-format` skill (`.opencode/skills/add-bank-format/`).

### Examples

Run `.venv\Scripts\python.exe examples/generate_examples.py` to create `examples/excel.xlsx` + `examples/index.html` with 100 synthetic transactions for testing without real data.

### Skills
- `setup` — first-time project initialization (uv, venv, pip install)
- `add-bank-format` — add a new PDF template to `config.yaml` + parser in `extract.py`
- `add-feature` — gitflow workflow for developing features
- `manage-categories` — add/edit/remove category rules in `config_personal.yaml` (or `config.yaml` as base)
- `categorize` — AI-assisted categorization: dump unknowns, fill categories, apply to XLSX
- `report` — run the full pipeline: extract → categorize (dump → AI → apply) → report

### Dependencies
- Python 3.14 (uv-managed), packages in `.venv/`
