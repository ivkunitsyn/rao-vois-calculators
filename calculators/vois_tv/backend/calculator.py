"""VOIS TV calculator entrypoint.

Структура шагов и общий расчётный пайплайн совпадают с РАО ТВ,
но ставки и переменные берутся из таблиц ВОИС.
"""

from calculators.rao_tv.backend.rao import run_calc_capture, parse_inn, get_org_name_by_inn, load_licenses_by_inn, fix_mojibake

__all__ = [
    "run_calc_capture",
    "parse_inn",
    "get_org_name_by_inn",
    "load_licenses_by_inn",
    "fix_mojibake",
]
