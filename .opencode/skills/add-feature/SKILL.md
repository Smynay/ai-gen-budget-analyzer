---
name: add-feature
description: Gitflow workflow for developing features in the budget repo
metadata:
  scope: "budget repo development process"
---

## Gitflow rules

### Branch naming
- Фичи: `feat/краткое-описание`
- Баги: `fix/краткое-описание`

### Workflow

1. **Создать ветку** от master:
   ```powershell
   git checkout master
   git checkout -b feat/название
   ```

2. **Разработка** — коммиты в ветку

3. **Пуш в GitHub** (без мержа):
   ```powershell
   git push origin feat/название
   ```

4. **Ревью** — автор показывает изменения, user проверяет

5. **Исправление замечаний** — коммиты в ту же ветку, пуш

6. **Ветка остаётся** — мерж в master только по явному запросу

### Правила

- **Никогда** не мержить свою ветку в master
- **Никогда** не форспушить в master
- Сквошить коммиты перед пушем при необходимости (`git reset --soft HEAD~N`)
- Перед пушем проверить: `extract.py` + `report.py` без ошибок
- Ветка остаётся на GitHub до указания user
