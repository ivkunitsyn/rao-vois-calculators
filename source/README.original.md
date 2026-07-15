# RNK — калькуляторы ставок (РАО/ВОИС)

## Что в репозитории

Репозиторий разделён на четыре уровня:

- `calculators/` — бизнес-логика каждого калькулятора.
- `interface/` — общий frontend-шелл и пользовательский интерфейс.
- `tools/` — сервисные утилиты, включая бот обновления таблицы РКН.
- `docs/` — эксплуатационная и аналитическая документация.

## Калькуляторы

Каждый калькулятор живёт в своей папке и имеет собственный Python-модуль `backend/calculator.py`.

- `calculators/rao_tv/` — РАО, вещатели телеканалов.
- `calculators/rao_radio/` — РАО, вещатели радиоканалов.
- `calculators/rao_events/` — РАО, трансляторы мероприятий.
- `calculators/vois_tv/` — ВОИС, вещатели телеканалов.
- `calculators/vois_radio/` — ВОИС, вещатели радиоканалов.
- `calculators/vois_events/` — ВОИС, трансляторы мероприятий.

Важно:

- для ТВ-калькуляторов API сейчас общее: `calculators/rao_tv/backend/api.py`;
- при этом точки входа по калькуляторам уже разделены и названы отдельно;
- РАО ТВ и ВОИС ТВ используют общий ТВ-движок расчёта, но разные таблицы ставок и разные режимы вызова;
- калькуляторы мероприятий используют отдельные Python-модули и не зависят от ТВ-логики;
- радио-калькуляторы пока оставлены как отдельные placeholder-модули, чтобы их можно было развивать независимо и без смешения логики ТВ.

Подробная карта структуры: `docs/PROJECT_STRUCTURE.md`.

## Запуск локально

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app:app --host 0.0.0.0 --port 8000
```

## Утилиты

```bash
# Очистка таблицы РКН
python -m tools.rkn_table_bot.rkn_cleaner --input ./raw.xlsx --output ./clean.xlsx --sqlite ./rkn.sqlite

# Telegram-бот обновления таблицы
python -m tools.rkn_table_bot.rkn_telegram_bot

# Аудит структуры проекта и проверка цепочки данных РКН
python -m tools.project_audit
```

## Деплой

CI/CD: `.gitlab-ci.yml` (build + deploy в docker).
