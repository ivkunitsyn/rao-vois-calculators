"""RAO TV calculator entrypoint.

Логика расчёта находится в `rao.py`.
Этот модуль выделен как отдельная точка входа для калькулятора РАО ТВ.
"""

from .rao import run_calc_capture, parse_inn, get_org_name_by_inn, load_licenses_by_inn, fix_mojibake

__all__ = [
    "run_calc_capture",
    "parse_inn",
    "get_org_name_by_inn",
    "load_licenses_by_inn",
    "fix_mojibake",
]
