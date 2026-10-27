# Выход ТС (КБМ)

PWA-сайт с ежедневной сводкой техники по цехам.

Репозиторий: `github.com/aura2889/summary` (перенесён из `u-seed/kbm-tehnika`).

## Как это работает

1. Outlook → скрипт PowerShell на ноутбуке → GitHub API кладёт файлы в `input/`.
2. GitHub Actions (`.github/workflows/build.yml`, `scripts/build.py`) собирает `data.json`, `drivers.json`, `archive/<дата>/`, `svod.xlsx`.
3. Сайт публикуется на GitHub Pages (ветка `main`, корень).
4. Копия отправляется на Bitbucket (шаг в `build.yml` и `mirror.yml`, секрет `BITBUCKET_TOKEN`).

Подробная ежедневная инструкция — в `ИНСТРУКЦИЯ.txt`.

## Структура

- `index.html`, `sw.js`, `manifest.json` — сайт (PWA)
- `data.json`, `drivers.json` — данные дня
- `archive/` — архив по датам (365 дней)
- `prod/` — добыча
- `scripts/` — сборка; правила цехов в `scripts/svod/parse_current.py`, план по цехам в `scripts/svod/resources/plan_dept.json`
- `input/` — сюда загружаются исходные файлы (после обработки удаляются)

## Локальная пересборка

```
SVOD_REBUILD=1 SVOD_DATE=ГГГГ-ММ-ДД python3 scripts/build.py
```

После полуночи `SVOD_DATE` указывать обязательно. Затем `git commit` и `git pull --rebase -X theirs && git push`.
