# Categorize Skill

AI-assisted categorization pipeline. After extracting PDFs, the AI reviews unknown transactions and assigns categories based on description semantics.

## Workflow

1. **Extract** — run `extract.py` to parse PDFs (keyword-based first pass, rest → `Неизвестно`)
2. **Dump** — run `categorize.py --dump` to group similar unknown descriptions into YAML
3. **AI review** — AI (you) reads the YAML, analyzes each group, and fills `category` based on meaning
4. **Apply** — run `categorize.py --apply` to write categories back to XLSX
5. **Report** — run `report.py` to generate HTML

## categorize.py

```
.venv\Scripts\python.exe categorize.py --dump
.venv\Scripts\python.exe categorize.py --apply
```

- `--dump`: reads latest XLSX, finds `Неизвестно` transactions, normalizes descriptions (strips prefixes, replaces numbers), groups similar ones, saves `results/YYYY-MM-DD_N/descriptions.yaml`
- `--apply` (no arg): reads `descriptions.yaml` from same directory as latest XLSX; `--apply PATH` for custom YAML path

## AI Task

For each group in `descriptions.yaml`:

1. Look at the sample `descriptions` and the sign of `total_amount` (negative = expense, positive = income)
2. Pick the best matching category from `config_personal.yaml` or `config.yaml`
3. Set `category` to the exact category name (use `Неизвестно` or leave `null` to skip)
4. Set `category` on individual group entries, not on the file level

When the AI is done with all groups, run `categorize.py --apply` → `report.py`.

## Notes

- Description normalization removes common prefixes like `ОПЛАТА В `, `БЕЗНАЛИЧНАЯ ОПЛАТА `, `ПРОЧИЕ ОПЕРАЦИИ `, replaces all numbers with `{N}`
- Transaction direction (+/-) should guide category choice (e.g., positive = Возврат долга, Доходы, Кешбеки)
- If unsure, set `category: Неизвестно` — better to leave uncategorized than mis-categorize
