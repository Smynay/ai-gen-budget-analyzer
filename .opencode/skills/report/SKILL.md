---
name: report
description: Full pipeline — extract, categorize, and report
metadata:
  scope: "budget repo full pipeline"
---

## What this skill does

Runs the complete personal finance pipeline from PDFs to interactive HTML report:

```
extract.py → categorize.py --dump → [AI reviews YAML] → categorize.py --apply → report.py
```

Steps:
1. **Extract** — parse PDFs, keyword-based first pass, rest → `Неизвестно`
2. **Dump unknowns** — group similar `Неизвестно` descriptions into YAML
3. **AI categorization** — agent fills categories in YAML based on description semantics
4. **Apply** — write categories back to XLSX
5. **Report** — generate interactive HTML

---

## Step 1: Extract

```powershell
.venv\Scripts\python.exe extract.py
```

Output: `results/YYYY-MM-DD_N/excel.xlsx`
- Сводный sheet (all transactions + Источник column) + one sheet per PDF (without Источник)
- Column order: Дата, Сумма, Валюта, Категория, Описание, (Источник)
- Keyword-based auto-categorization from `config_personal.yaml` → categories (or `config.yaml` as fallback)
- Optional `direction` field (`income` / `expense`) limits matching by transaction sign
- Unmatched → `Неизвестно` (fallback)

---

## Step 2: Dump unknowns

```powershell
.venv\Scripts\python.exe categorize.py --dump
```

Output: `results/YYYY-MM-DD_N/descriptions.yaml`

Groups `Неизвестно` transactions by normalized description (strips known prefixes, replaces numbers with `{N}`, uppercase). Each group shows count, sample descriptions, and total amount.

---

## Step 3: AI categorization

Agent reads `descriptions.yaml` and fills `category` for each group:

- **Look at the descriptions** — identify the merchant or purpose
- **Check `total_amount` sign** — negative = expense, positive = income
- **Pick the best category** from available categories in `config_personal.yaml`
- **Set `category`** field on the group level (not file level)
- **If unsure**, set `null` or `Неизвестно` — better to skip than mis-categorize

Sets `category: Неизвестно` or leaves `null` for groups the agent cannot confidently categorize.

---

## Step 4: Apply categories

```powershell
.venv\Scripts\python.exe categorize.py --apply
```

Reads `descriptions.yaml` from the same directory as the latest XLSX and updates `Категория` column in the XLSX for all matched transactions.

---

## Step 5: Generate HTML report

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

---

## Quick start (no unknowns or already categorized)

If no AI categorization is needed, skip steps 2–4:

```powershell
.venv\Scripts\python.exe extract.py
.venv\Scripts\python.exe report.py
```

---

## Troubleshooting

| Log message | Meaning | Fix |
|---|---|---|
| `No matching template for X.pdf` | Unknown bank format | Use `add-bank-format` skill |
| `No results directories found` | `extract.py` hasn't been run yet | Run `extract.py` first |
| `excel.xlsx not found` | Directory exists but missing XLSX | Re-run `extract.py` |
| `Неизвестно` entries in report | Some transactions weren't categorized | Run `categorize.py --dump`, let AI fill YAML, then `categorize.py --apply` |
| `descriptions.yaml not found` | `--dump` hasn't been run | Run `categorize.py --dump` first |
| Console shows `?` for Russian text | Terminal encoding (cp1251) | Output files (XLSX/HTML) are UTF-8 — check by opening them |
