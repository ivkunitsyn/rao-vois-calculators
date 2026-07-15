from __future__ import annotations

import base64
import html
import json
import re
from typing import Any

from .formatting import format_internet_component_label, format_money, format_number, format_percent
from .model import ContractTerms, MinimumLicense, RadioReportModel, ReportRow

MODEL_MARKER_PREFIX = "RADIO_REPORT_MODEL:"


def _h(value: Any) -> str:
    return html.escape("" if value is None else str(value), quote=True)


def _model_comment(model: RadioReportModel) -> str:
    payload = json.dumps(model.to_dict(), ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    encoded = base64.b64encode(payload).decode("ascii")
    return f"<!--{MODEL_MARKER_PREFIX}{encoded}-->"


def extract_model_from_html(raw_html: str) -> RadioReportModel | None:
    match = re.search(rf"<!--{re.escape(MODEL_MARKER_PREFIX)}([A-Za-z0-9+/=]+)-->", str(raw_html or ""))
    if not match:
        return None
    try:
        data = json.loads(base64.b64decode(match.group(1)).decode("utf-8"))
        return RadioReportModel.from_dict(data)
    except Exception:
        return None


def _license_link(number: str, url: str) -> str:
    value = str(number or "").strip()
    if not value:
        return "—"
    return (
        f'<a class="radioReportLicenseLink" href="{_h(url)}" '
        f'target="_blank" rel="noopener noreferrer">{_h(value)}</a>'
    )


def _value(row: ReportRow) -> str:
    value = row.value
    if value is None:
        return "—"
    return str(value)


def _facts(rows: list[ReportRow]) -> str:
    body: list[str] = []
    for row in rows:
        value = _value(row)
        if value == "":
            continue
        marker = " radioReportResultFact" if row.key_result and value.strip() != "—" else ""
        body.append(
            f'<p class="radioReportFact{marker}">'
            f'<strong>{_h(row.label)}:</strong> {_h(value)}'
            "</p>"
        )
    return '<div class="radioReportFacts">' + "".join(body) + "</div>"


def _section(title: str, body: str) -> str:
    return (
        '<section class="radioReportSection">'
        f"<h3>{_h(title)}</h3>"
        f"{body}"
        "</section>"
    )


def _rate_licenses(model: RadioReportModel) -> str:
    blocks: list[str] = []
    for lic in model.rate_licenses:
        heading = '<span class="radioReportMarker">Новая лицензия</span>' if lic.is_new else "Лицензия"
        blocks.append('<div class="radioReportLicenseBlock">')
        blocks.append(
            '<div class="radioReportLicenseTitle">'
            f'{heading} № {_license_link(lic.license_number, lic.license_url)} — {format_percent(lic.rate)}'
            "</div>"
        )
        if not lic.channels:
            blocks.append('<div class="radioReportMuted">СМИ/радиоканалы не указаны.</div>')
        for ch in lic.channels:
            blocks.append('<div class="radioReportChannel">')
            blocks.append(
                f'<div class="radioReportChannelTitle">{_h(ch.name)} — '
                f'{format_number(ch.weekly_hours)} ч — {format_percent(ch.rate)}</div>'
            )
            if ch.actual_share_percent is not None:
                blocks.append(
                    '<div class="radioReportMuted">Фактическая доля использования: '
                    f'{format_percent(ch.actual_share_percent)}. Ставка: {format_percent(ch.rate)}.</div>'
                )
            elif ch.topics:
                blocks.append('<ul class="radioReportList">')
                for topic in ch.topics:
                    blocks.append(
                        f"<li>{_h(topic.name)} → категория {_h(topic.category)} — "
                        f"{format_percent(topic.rate)}</li>"
                    )
                blocks.append("</ul>")
            blocks.append("</div>")
        blocks.append("</div>")
    return "".join(blocks)


def _hours_coefficient_label(value: float | None) -> str:
    if value is None:
        return "—"
    return "не применяется" if abs(float(value) - 1.0) < 1e-9 else str(value).replace(".", ",")


def _minimum_license(lic: MinimumLicense, trailing_rows: list[ReportRow] | None = None) -> str:
    heading = '<span class="radioReportMarker">Новая лицензия</span>' if lic.is_new else "Лицензия"
    internet_label = format_internet_component_label(lic.internet_resources)
    internet_line = (
        f'<div class="radioReportLicenseAddon">{_h(internet_label)}: '
        f'{format_money(lic.internet_component)} ₽.</div>'
        if float(lic.internet_component or 0.0) > 0
        else ""
    )
    rows = [
        ReportRow("Численность населения", format_number(lic.population)),
        ReportRow("Диапазон по численности населения", lic.population_range or "—"),
        ReportRow("Интернет-ресурсы", format_number(lic.internet_resources)),
        ReportRow("Всего часов вещания", format_number(lic.weekly_hours)),
        ReportRow("Коэффициент за объём вещания", _hours_coefficient_label(lic.hours_coefficient)),
    ]
    return (
        '<div class="radioReportLicenseBlock">'
        f'<div class="radioReportLicenseTitle radioReportMinimumTitle">{heading} № '
        f'{_license_link(lic.license_number, lic.license_url)} — минимальная сумма: '
        f'{format_money(lic.minimum)} ₽</div>'
        f"{internet_line}"
        f"{_facts(rows + list(trailing_rows or []))}"
        "</div>"
    )


def _contract_terms(terms: ContractTerms) -> str:
    parts = ['<div class="radioReportTermsBlock">', f"<h4>{_h(terms.title)}</h4>"]
    if not terms.lines and not terms.bullets:
        parts.append("<p>—</p>")
    for line in terms.lines:
        parts.append(f"<p>{_h(line)}</p>")
    if terms.bullets:
        parts.append('<ul class="radioReportList">')
        for bullet in terms.bullets:
            parts.append(f"<li>{_h(bullet)}</li>")
        parts.append("</ul>")
    parts.append("</div>")
    return "".join(parts)


def render_radio_report_html(model: RadioReportModel, *, embed_model: bool = True) -> str:
    sections = [
        _section("1. ИСХОДНЫЕ ДАННЫЕ", _facts(model.source_data)),
        _section(
            "2. РАСЧЁТ ПРОЦЕНТНОЙ СТАВКИ",
            _rate_licenses(model)
            + _facts(
                [
                    ReportRow("Процентная ставка по договору", format_percent(model.contract_rate), True),
                    ReportRow(
                        "Расчётная сумма за квартал",
                        f"{format_money(model.quarter_amount, precise=True)} ₽" if model.quarter_amount is not None else "—",
                    ),
                ]
            ),
        ),
        _section(
            "3. РАСЧЁТ МИНИМАЛЬНОЙ СУММЫ ЗА КВАРТАЛ",
            (
                "".join(
                    _minimum_license(
                        x,
                        trailing_rows=model.minimum_rows if idx == len(model.minimum_licenses) - 1 else None,
                    )
                    for idx, x in enumerate(model.minimum_licenses)
                )
                if model.minimum_licenses
                else _facts(model.minimum_rows)
            ),
        ),
        _section("4. УСЛОВИЯ ДОГОВОРА", "".join(_contract_terms(x) for x in model.contract_terms)),
        _section(
            "5. КОММЕНТАРИИ К РАСЧЁТУ",
            '<ul class="radioReportList">'
            + "".join(f"<li>{_h(x)}</li>" for x in model.comments)
            + "</ul>"
            if model.comments
            else "<p>Дополнительных комментариев нет.</p>",
        ),
    ]
    html_out = '<article class="radioReport"><h1>Результат расчёта</h1>' + "".join(sections) + "</article>"
    return html_out + (_model_comment(model) if embed_model else "")
