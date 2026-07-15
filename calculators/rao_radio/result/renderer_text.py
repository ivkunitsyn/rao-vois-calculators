from __future__ import annotations

from .formatting import format_internet_component_label, format_money, format_number, format_percent
from .model import ContractTerms, MinimumLicense, RadioReportModel, ReportRow


def _row(row: ReportRow) -> str:
    return f"{row.label}: {row.value if row.value is not None else '—'}"


def _terms(block: ContractTerms) -> list[str]:
    lines = [block.title]
    if not block.lines and not block.bullets:
        lines.append("—")
    lines.extend(block.lines)
    lines.extend(f"• {x}" for x in block.bullets)
    return lines


def _minimum_license(lic: MinimumLicense, trailing_rows: list[ReportRow] | None = None) -> list[str]:
    prefix = "Новая лицензия" if lic.is_new else "Лицензия"
    lines = [
        f"{prefix} № {lic.license_number} — минимальная сумма: {format_money(lic.minimum)} ₽",
    ]
    if float(lic.internet_component or 0.0) > 0:
        label = format_internet_component_label(lic.internet_resources)
        lines.append(f"{label}: {format_money(lic.internet_component)} ₽.")
    coeff = "не применяется" if lic.hours_coefficient is None or abs(float(lic.hours_coefficient) - 1.0) < 1e-9 else str(lic.hours_coefficient).replace(".", ",")
    lines.extend(
        [
            f"Численность населения: {format_number(lic.population)}",
            f"Диапазон по численности населения: {lic.population_range or '—'}",
            f"Интернет-ресурсы: {format_number(lic.internet_resources)}",
            f"Всего часов вещания: {format_number(lic.weekly_hours)}",
            f"Коэффициент за объём вещания: {coeff}",
        ]
    )
    lines.extend(_row(row) for row in trailing_rows or [])
    return lines


def render_radio_report_text(model: RadioReportModel) -> str:
    lines: list[str] = ["Результат расчёта", ""]
    lines.append("1. ИСХОДНЫЕ ДАННЫЕ")
    lines.extend(_row(x) for x in model.source_data)
    lines.append("")

    lines.append("2. РАСЧЁТ ПРОЦЕНТНОЙ СТАВКИ")
    for lic in model.rate_licenses:
        prefix = "Новая лицензия" if lic.is_new else "Лицензия"
        lines.append(f"{prefix} № {lic.license_number} — {format_percent(lic.rate)}")
        for ch in lic.channels:
            lines.append(f"{ch.name} — {format_number(ch.weekly_hours)} ч — {format_percent(ch.rate)}")
            if ch.actual_share_percent is not None:
                lines.append(f"Фактическая доля использования: {format_percent(ch.actual_share_percent)}. Ставка: {format_percent(ch.rate)}.")
            else:
                lines.extend(f"• {t.name} → категория {t.category} — {format_percent(t.rate)}" for t in ch.topics)
    lines.append(f"Процентная ставка по договору: {format_percent(model.contract_rate)}")
    lines.append(
        "Расчётная сумма за квартал: "
        + (f"{format_money(model.quarter_amount, precise=True)} ₽" if model.quarter_amount is not None else "—")
    )
    lines.append("")

    lines.append("3. РАСЧЁТ МИНИМАЛЬНОЙ СУММЫ ЗА КВАРТАЛ")
    if model.minimum_licenses:
        last_idx = len(model.minimum_licenses) - 1
        for idx, lic in enumerate(model.minimum_licenses):
            lines.extend(_minimum_license(lic, trailing_rows=model.minimum_rows if idx == last_idx else None))
    else:
        lines.extend(_row(x) for x in model.minimum_rows)
    lines.append("")

    lines.append("4. УСЛОВИЯ ДОГОВОРА")
    for block in model.contract_terms:
        lines.extend(_terms(block))
    lines.append("")

    lines.append("5. КОММЕНТАРИИ К РАСЧЁТУ")
    if model.comments:
        lines.extend(f"• {x}" for x in model.comments)
    else:
        lines.append("Дополнительных комментариев нет.")
    return "\n".join(lines).strip()
