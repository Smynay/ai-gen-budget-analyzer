---
name: report
description: Build a spending/income report from the extracted XLSX
metadata:
  scope: "budget repo result generation"
---

## What this skill does

Two pipelines:

1. **Extraction**: `extract.py` — reads PDFs from `sources/`, writes structured output to `results/YYYY-MM-DD_N/`
2. **HTML report**: `report.py` — reads the latest XLSX, generates an interactive HTML report with Chart.js pie chart, category table, and transaction drill-down

## Step 1: Run extraction

```powershell
.venv\Scripts\python.exe extract.py
```

Output goes to `results/YYYY-MM-DD_N/`:
- `excel.xlsx` — Сводный sheet (all transactions + Источник column) + one sheet per PDF (without Источник)
- Column order: Дата, Сумма, Валюта, Категория, Описание, (Источник)

## Step 2: Generate HTML report

```powershell
.venv\Scripts\python.exe report.py                    # весь датасет
.venv\Scripts\python.exe report.py --year 2026         # только 2026 год
.venv\Scripts\python.exe report.py --month 2026-01     # январь 2026
.venv\Scripts\python.exe report.py --from 2026-01-01 --to 2026-06-30  # произвольный период
.venv\Scripts\python.exe report.py --from 2026-01-01   # с даты до конца
.venv\Scripts\python.exe report.py --to 2026-06-30     # с начала до даты
```

Output: `results/YYYY-MM-DD_N/report.html` (или `report-2026.html`, `report-2026-01.html` и т.д.)

Open in any browser. Features:
- **Toggle**: Расходы / Доходы / Всё — pie chart updates live
- **Pie chart**: each slice = category with absolute total + percentage tooltip — click to drill down
- **Category table**: name, total, count — click any row to drill down
- **Drill-down**: shows date (sortable), amount, description, source (toggleable) for every transaction in the category
- **Source checkbox**: show/hide the Источник column in drill-down
- **Summary cards**: total transactions, income, expenses, net

## Category system

Transactions are auto-categorized by keywords (defined in `config_personal.yaml` → `categories`, falling back to `config.yaml`). Categories can optionally specify a `direction` (`income` / `expense`) to only match transactions with the corresponding sign. The `Категория` column is generated during extraction and available in the report.

## Troubleshooting

| Log message | Meaning | Fix |
|---|---|---|
| `No matching template for X.pdf` | Unknown bank format | Use `add-bank-format` skill |
| `No results directories found` | `extract.py` hasn't been run yet | Run `extract.py` first |
| `excel.xlsx not found` | Directory exists but missing XLSX | Re-run `extract.py` |
| Console shows `?` for Russian text | Terminal encoding (cp1251) | Output files (XLSX/HTML) are UTF-8 — check by opening them |
