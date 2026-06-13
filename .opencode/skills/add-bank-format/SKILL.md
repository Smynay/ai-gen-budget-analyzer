---
name: add-bank-format
description: Add a new bank PDF template to config.yaml and a parser to extract.py
metadata:
  scope: "budget repo extraction pipeline"
---

## What to do

A new bank statement PDF was placed in `sources/` and doesn't match existing templates. You need to add support for it.

## Step 1: Investigate the PDF

```python
import fitz
from collections import defaultdict

doc = fitz.open("sources/your_file.pdf")
# check metadata for matching keywords
print(doc.metadata)

# check first-page text
print(doc[0].get_text()[:500])

# get word positions to find column layout
words = doc[0].get_text("words")
rows = defaultdict(list)
for w in words:
    y_key = round(w[1] / 3) * 3
    rows[y_key].append(w)
for y in sorted(rows):
    r = sorted(rows[y], key=lambda w: w[0])
    print(f"y={y}: " + " | ".join(f"{w[4]}[x={w[0]:.0f}]" for w in r))
doc.close()
```

This reveals:
- **Keywords** to use in `match` (from metadata or first-page text)
- **Column x-ranges**: where each column's words fall (date, amount, description)
- **Amount format**: decimal separator (`.` or `,`), group separator (space or none), sign convention

## Step 2: Add template to `config.yaml`

```yaml
templates:
  - name: "your_bank_type"       # snake_case, used in code routing
    match:
      keywords: ["Unique bank name from metadata/text"]
    amount:
      decimal_separator: "."
      group_separator: " "
    columns:
      date:        { x_start: 30,  x_end: 125 }
      description: { x_start: 140, x_end: 360 }
      amount:      { x_start: 360, x_end: 510 }
```

Name columns however your layout defines them (e.g. `date_op`, `date_debit`, `amount_txn`, `amount_card` for multi-column layouts). The parser function will reference these keys.

## Step 3: Add parser to `extract.py`

Create a function following this pattern. Common patterns to adapt:

### Continuation / multi-row descriptions
If a transaction description can span multiple rows (e.g. Tinkoff), track `desc_lines` and merge:

```python
        txn = None
        desc_lines = []

        for y, row_words in rows:
            amt_text = column_text(...)
            has_amount = is_numeric_value(amt_text)

            if has_amount:
                if txn:
                    merged_desc = (txn.get("description", "") + " " + " ".join(desc_lines)).strip()
                    if merged_desc:
                        txn["description"] = merged_desc
                    if txn["date"] and txn["amount"] is not None:
                        transactions.append(txn)

                date_text = column_text(...)
                parsed_date = parse_russian_date(date_text)
                amount = parse_russian_amount(...)
                desc = column_text(...)

                txn = {"date": parsed_date, "amount": amount, "description": desc}
                desc_lines = []
            elif txn and len(desc_lines) < 4:
                desc_text = column_text(...)
                if desc_text:
                    desc_lines.append(desc_text)

        if txn:
            merged_desc = (txn.get("description", "") + " " + " ".join(desc_lines)).strip()
            if merged_desc:
                txn["description"] = merged_desc
            if txn["date"] and txn["amount"] is not None:
                transactions.append(txn)
```

Remove duplicate words in descriptions:
```python
    merged = []
    for t in transactions:
        if t["date"] is not None and t["amount"] is not None:
            if t["description"]:
                parts = t["description"].split()
                seen = set()
                deduped = []
                for p in parts:
                    if p not in seen:
                        seen.add(p)
                        deduped.append(p)
                t["description"] = " ".join(deduped)
            merged.append(t)
    return merged
```

### Optional: date+time parsing
If amounts start on the same row as date+time (e.g. Sberbank), use `parse_russian_datetime()` instead of `parse_russian_date()`.

### Footer / header text filtering
If each page has a "Продолжение на следующей странице" footer, filter it during continuation row handling:
```python
elif txn and len(desc_lines) < 4:
    if cat_text and any(kw in cat_text.upper() for kw in ("ПРОДОЛЖЕНИЕ", "СЛЕДУЮЩЕЙ", "СТРАНИЦЕ")):
        continue
    ...
```

Then register it in `extract_transactions()` at the bottom of the file:

```python
def extract_transactions(pdf_path, template):
    name = template["name"]
    if name == "sberbank_debit":
        return extract_sberbank(pdf_path, template)
    elif name == "tinkoff_movement":
        return extract_tinkoff(pdf_path, template)
    elif name == "your_bank_type":          # <-- add this
        return extract_your_bank(pdf_path, template)
    raise ValueError(f"Unknown template: {name}")
```

### Key helpers available

| Helper | Purpose |
|--------|---------|
| `group_words_by_row(words)` | Group PyMuPDF word tuples by y-coordinate |
| `column_text(words, x_start, x_end)` | Get combined text from words within an x-range |
| `is_numeric_value(text)` | Check if text is a parseable number |
| `parse_russian_amount(text, dec, grp)` | Parse "1 234,56" or "-1 234.56" to float |
| `parse_russian_date(text)` | Extract "dd.mm.yyyy" from text → datetime |
| `find_header_row(rows, cols, text)` | Find table header row index |

### Sign convention patterns

- **Tinkoff style**: `-880.00` = expense, `+5 000.00` = income → use amount as-is
- **Sberbank style**: unsigned `1 234,56` = expense, `+400,00` = income → negate unsigned amounts
- Document the convention in a comment alongside the amount parsing line

## Step 4: Test

```powershell
.venv\Scripts\python.exe extract.py
```

Check `results/YYYY-MM-DD_N/excel.xlsx` — verify dates parsed correctly, amounts have correct signs, descriptions aren't truncated, and no header/footer text leaked in.
