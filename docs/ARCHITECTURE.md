# RNK Architecture

## Source Of Truth

- Authoritative code remote/branch: GitLab `sonic/main`.
- Deprecated remote: GitHub `origin`; do not use it for development.

## Runtime Entrypoint

`app.py` imports the FastAPI application from:

`calculators/rao_tv/backend/api.py`

Despite the `rao_tv` path, this API module is shared by all calculators.

## Frontend

`interface/index.html` is the current frontend monolith. It contains:

- calculator routing;
- wizard UI;
- form/session state;
- result display;
- admin panel UI;
- copy actions;
- DOCX export requests.

## Backend

`calculators/rao_tv/backend/api.py` currently owns:

- FastAPI app;
- static routes;
- calculation dispatch;
- registry/license endpoints;
- DOCX export;
- RKN admin upload/status endpoints;
- usage statistics.

## Engines

| Calculator | Runtime path |
| --- | --- |
| TV RAO | `calculators/rao_tv/backend/rao.py` |
| TV VOIS | `calculators/vois_tv/backend/` wrapper over TV RAO family |
| RV RAO | `calculators/rao_radio/backend/rao_radio.py` |
| RV VOIS | `calculators/vois_radio/backend/` wrapper over RV RAO family |
| Online RAO | `calculators/rao_events/backend/calculator.py` |
| Online VOIS | `calculators/vois_events/backend/` wrapper over Online RAO family |

## Dependency Warnings

- Radio code imports TV RAO helpers for registry/parsing behavior.
- VOIS calculators reuse RAO engines through wrappers.
- Registry tooling imports helpers from the TV RAO engine.
- `calculators/rao_tv/backend/api.py` affects all calculators.
- `interface/index.html` affects all calculators.

## Runtime Data

Current registry-related artifacts include:

- `calculators/rao_tv/data/Таблица РКН.xlsx`;
- `calculators/rao_tv/data/Таблица РКН slim.xlsx`;
- `calculators/rao_tv/data/Таблица РКН.sqlite`;
- server `/data` overrides and runtime upload logs.

CURRENT STATE: multiple possible registry artifacts exist. The final runtime source of truth still needs a human decision.
