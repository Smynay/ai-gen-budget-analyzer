---
name: manage-categories
description: Add, remove, or edit auto-categorization rules in config.yaml or config_personal.yaml
metadata:
  scope: "budget repo category management"
---

## What this skill does

The `categories` section defines keyword-based rules that auto-categorize every transaction when `extract.py` runs. Categories live in either:
- **`config.yaml`** — базовые/пример категорий (коммитится в git)
- **`config_personal.yaml`** — твои персональные категории (не коммитится, добавлен в `.gitignore`)

Если `config_personal.yaml` существует — `extract.py` использует его секцию `categories`. Если нет — использует `categories` из `config.yaml`.

## Structure

```yaml
categories:
  fallback: "Другое"              # default category when no rule matches
  items:
    - name: "Супермаркеты"        # category name (appears in Категория column)
      direction: expense           # optional: "income" / "expense" / omit (any)
      keywords:                    # list of uppercase substrings to match
        - "ПЯТЕРОЧКА"
        - "LENTA"
```

- Matching is case-insensitive (descriptions are uppercased before matching)
- Categories are checked **in order** — first match wins
- Put more specific keywords/categories before broad ones
- `direction` limits matching by transaction sign: `income` (amount > 0), `expense` (amount < 0), or omit for any sign

## Common tasks

### Add a new category

Add a new `- name: ...` block under `categories.items` in your personal config:

```yaml
    - name: "Медицина"
      keywords:
        - "АПТЕКА"
        - "ПОЛИКЛИНИКА"
        - "БОЛЬНИЦА"
```

To limit by transaction type, add `direction`:
```yaml
    - name: "Возврат долга"
      direction: income          # only matches incoming transactions (amount > 0)
      keywords:
        - "Б. МАКСИМ НИКОЛАЕВИЧ"
```

### Add keywords to an existing category

Add one or more lines to `keywords:`:

```yaml
    - name: "Супермаркеты"
      keywords:
        - ...
        - "SPAR"          # added
        - "АШАН"          # added
```

### Remove a category or keyword

Just delete the relevant lines from the file you're editing (`config_personal.yaml` or `config.yaml`).

### Change the fallback

Change `fallback:` value (default: `"Другое"`).

## Verify

After any change, re-run and check the output:

```powershell
.venv\Scripts\python.exe extract.py
```

Spot-check a few transactions per category:

```python
from openpyxl import load_workbook
from collections import Counter
import glob
xlsx = sorted(glob.glob("results/**/excel.xlsx", recursive=True))[-1]
wb = load_workbook(xlsx)
ws = wb["Сводный"]
cats = Counter()
for row in ws.iter_rows(min_row=2, values_only=True):
    cats[row[3]] += 1  # column index of Категория
for cat, cnt in cats.most_common():
    print(f"{cat}: {cnt}")
```

## Tips

- Keywords are matched as **substrings** — `"ПЯТЕРОЧКА"` matches any description containing that word
- Short keywords (like `"GO"`) may produce false positives — prefer longer, more specific substrings
- Review the actual transaction descriptions before choosing keywords:

```python
from openpyxl import load_workbook
import glob
xlsx = sorted(glob.glob("results/**/excel.xlsx", recursive=True))[-1]
wb = load_workbook(xlsx)
ws = wb["Сводный"]
descs = set()
for row in ws.iter_rows(min_row=2, values_only=True):
    descs.add(row[4])  # column index of Описание
for d in sorted(descs):
    print(d)
```
