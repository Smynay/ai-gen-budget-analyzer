# budget

Извлекает транзакции из PDF-выписок российских банков и строит интерактивные отчёты.

## Быстрый старт

```powershell
uv venv
uv pip install PyMuPDF openpyxl pyyaml
```

## Использование

Положи PDF-выписки в `sources/`, затем:

```powershell
.venv\Scripts\python.exe extract.py
.venv\Scripts\python.exe report.py
```

Результат — `results/YYYY-MM-DD_N/report.html` (открыть в браузере).

## Пример данных

```powershell
.venv\Scripts\python.exe examples/generate_examples.py
```

Создаёт `examples/excel.xlsx` и `examples/index.html` — 100 синтетических транзакций по всем категориям. Можно открыть и посмотреть, как выглядит отчёт, без загрузки реальных выписок.

## Фильтры отчёта

```powershell
.venv\Scripts\python.exe report.py --year 2026
.venv\Scripts\python.exe report.py --month 2026-01
.venv\Scripts\python.exe report.py --from 2026-01-01 --to 2026-06-30
```

## Поддерживаемые банки

- **Сбербанк** — выписка по счёту дебетовой карты
- **Тинькофф** — справка о движении средств

## Настройка категорий

В `config_personal.yaml` можно добавлять и менять правила автокатегоризации (если его нет — используются категории из `config.yaml`):

```yaml
categories:
  fallback: "Другое"
  items:
    - name: "Кафе"
      keywords: ["РЕСТОРАН", "COFFEE"]

    - name: "Возврат долга"
      direction: income
      keywords: ["АНОН АНОНОВИЧ"]
```

`config_personal.yaml` не отслеживается git'ом. Чтобы вернуться к базовым категориям — просто удали его.
