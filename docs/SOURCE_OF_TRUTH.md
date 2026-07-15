# RNK Source Of Truth

## Code

Authoritative:

- GitLab `sonic/main`

Deprecated:

- GitHub `origin`

`origin` is historical/outdated and must not be used for development, fallback, or baseline comparisons.

## Registry Data

Existing registry artifacts:

- `calculators/rao_tv/data/Таблица РКН.xlsx`;
- `calculators/rao_tv/data/Таблица РКН slim.xlsx`;
- `calculators/rao_tv/data/Таблица РКН.sqlite`;
- server `/data` runtime overrides;
- RKN upload logs and current-file metadata.

CURRENT STATE: multiple possible registry artifacts exist.

DECISION REQUIRED: choose one runtime source of truth for the registry and document how admin uploads, sqlite indexes, and tracked xlsx files relate to it.

## Rates

- TV RAO: `calculators/rao_tv/data/Переменные из ставок.xlsx` plus code constants/helpers in `calculators/rao_tv/backend/rao.py`.
- TV VOIS: `calculators/vois_tv/data/Переменные из ставок ВОИС.xlsx`, with wrapper reuse of TV RAO family code.
- Radio RAO/VOIS: constants and calculation tables in `calculators/rao_radio/backend/rao_radio.py`, with VOIS radio wrapper reuse.
- Events RAO/VOIS: profiles/constants in `calculators/rao_events/backend/calculator.py` and VOIS wrapper profile in `calculators/vois_events/backend/calculator.py`.

Some sources are currently mixed between xlsx files, constants, and wrappers.

## Result Rendering

Current presentation paths:

- calculation engines produce text and, for radio, structured/report HTML;
- frontend displays returned HTML/text in `interface/index.html`;
- DOCX export is implemented in `calculators/rao_tv/backend/api.py`;
- copy behavior is implemented in `interface/index.html`.

CURRENT STATE: multiple presentation paths exist. HTML, copy, and DOCX can diverge unless changed carefully.
