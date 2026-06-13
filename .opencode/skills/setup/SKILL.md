---
name: setup
description: Initialize the budget project from scratch — install Python, deps
metadata:
  scope: "budget repo first-time setup"
---

## What to do

This repo extracts transactions from Russian bank statement PDFs (`sources/*.pdf`) into structured output (`results/YYYY-MM-DD_N/excel.xlsx`) and generates interactive HTML reports.

Run once when cloning the repo fresh or restoring the environment.

## Prerequisites

- Windows (paths below are Windows-specific)
- Internet connection (for downloads)

## Step 1: Install Python (via uv)

If `uv` is not available, install it:

```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

Then let uv manage Python:

```powershell
uv python install 3.14
```

Verify:
```powershell
uv python list
```

## Step 2: Create virtual env and install packages

```powershell
uv venv
uv pip install PyMuPDF openpyxl pyyaml
```

This creates `.venv\` in the project root. The extraction script uses `.venv\Scripts\python.exe`.

## Step 3: Verify setup

```powershell
.venv\Scripts\python.exe extract.py
.venv\Scripts\python.exe report.py
```

Expected output:
- `results/YYYY-MM-DD_N/excel.xlsx` — all transactions extracted
- `results/YYYY-MM-DD_N/report.html` — interactive HTML report (open in browser)

For each PDF, logs:
- Template matched (`sberbank_debit`, `tinkoff_movement`, or custom)
- Transaction count extracted

## Configuration

- `config.yaml` — defines schema columns and PDF templates
- `config_personal.yaml` — personal category overrides (not tracked by git, see `.gitignore`)
- `.gitignore` — prevents `config_personal.yaml`, `results/`, `.venv/` from being committed
- `.venv\` — virtual environment (do not commit)
- `results\` — output directory (do not commit)

## Troubleshooting

| Problem | Likely fix |
|---------|-----------|
| `pip install` fails with "externally managed" | Use `uv pip install` instead of `pip` |
| `uv venv` fails | `uv` not installed, install it first |
| No template matches | Add a new template via `add-bank-format` skill |
| Garbled console output | Console encoding issue, the XLSX file is unaffected |
