from __future__ import annotations

from decimal import Decimal
from typing import Any, Optional
from urllib.parse import quote


def format_number(value: Optional[Any]) -> str:
    if value is None:
        return "—"
    d = Decimal(str(value))
    sign = "-" if d < 0 else ""
    d = abs(d)
    if d == d.to_integral_value():
        return f"{sign}{int(d):,}".replace(",", " ")
    raw = format(d.normalize(), "f")
    int_part, dot, frac_part = raw.partition(".")
    int_fmt = f"{int(int_part or '0'):,}".replace(",", " ")
    frac_part = frac_part.rstrip("0")
    return f"{sign}{int_fmt}{('.' + frac_part) if dot and frac_part else ''}"


def format_money(value: Optional[Any], *, precise: bool = False) -> str:
    if value is None:
        return "—"
    d = Decimal(str(value))
    if not precise:
        return f"{float(d):,.2f}".replace(",", " ").replace(".00", "")
    raw = format(d, "f")
    if "." in raw:
        raw = raw.rstrip("0").rstrip(".")
    sign = "-" if raw.startswith("-") else ""
    if sign:
        raw = raw[1:]
    int_part, dot, frac_part = raw.partition(".")
    int_fmt = f"{int(int_part or '0'):,}".replace(",", " ")
    return f"{sign}{int_fmt}{('.' + frac_part) if dot and frac_part else ''}"


def format_percent(value: Optional[Any]) -> str:
    if value is None:
        return "—"
    raw = f"{float(value):.1f}".replace(".", ",")
    if raw.endswith(",0"):
        raw = raw[:-2]
    return f"{raw}%"


def format_license_url(license_number: str) -> str:
    return "https://rkn.gov.ru/activity/mass-media/for-broadcasters/teleradio/?id=" + quote(
        str(license_number or ""),
        safe="",
    )


def format_internet_component_label(internet_resources: Optional[Any]) -> str:
    try:
        count = int(internet_resources or 0)
    except Exception:
        count = 0
    if count == 1:
        return "Включает сумму за сайт / приложение в размере"
    return "Включает сумму за интернет-ресурсы в размере"


def format_financial_base_type(value: Optional[str]) -> str:
    code = str(value or "").strip()
    if code == "annual_revenue":
        return "годовая выручка или доход"
    if code == "revenue_q":
        return "доходы за квартал"
    if code == "annual_expenses":
        return "расходы за год"
    if code == "expenses_q":
        return "расходы за квартал"
    return "не указана"


def format_internet_income_threshold_label(previous_year_income: Optional[Any]) -> str:
    amount = float(previous_year_income or 0.0)
    if amount <= 3_000_000:
        return "Доход до 3 000 000 ₽"
    if amount <= 50_000_000:
        return "Доход свыше 3 000 000 ₽ до 50 000 000 ₽"
    return "Доход свыше 50 000 000 ₽"
