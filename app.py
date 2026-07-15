"""Явная backend-точка входа для сервера.

Запуск:
    uvicorn app:app --host 0.0.0.0 --port 8000
"""

from calculators.rao_tv.backend.api import app

__all__ = ["app"]
