"""RAO events calculator engine.

Калькулятор ставок для онлайн-трансляторов мероприятий (РАО/ВОИС).
Логика покрывает ветки разделов 2-4 положений:
- стандартная категория;
- специальная 3.2.1 и 3.2.2;
- коэффициенты (бесплатный доступ, доля произведений/фонограмм);
- учет площадок и количества трансляций.
"""

from __future__ import annotations

from dataclasses import dataclass
import math
from typing import Any


CALCULATOR_ID = "rao_events"
PROVIDER = "РАО"


EVENT_RATES_TABLE: dict[str, dict[int, int]] = {
    "10000": {4: 12500, 8: 9374, 12: 8062, 24: 7458},
    "100000": {4: 23500, 8: 17624, 12: 15158, 24: 14020},
    "1000000": {4: 43000, 8: 32250, 12: 27734, 24: 25654},
    "unlimited": {4: 85000, 8: 63750, 12: 54824, 24: 50712},
}

EVENT_MUSIC_SHARE_COEFFICIENTS: dict[str, float] = {
    "19": 0.6,
    "39": 0.7,
    "59": 0.8,
    "79": 0.9,
    "100": 1.0,
}

EVENT_REGION_RATES: dict[str, int] = {
    "Белгородская область": 3229, "Брянская область": 2748, "Владимирская область": 6242, "Воронежская область": 3645,
    "Ивановская область": 2256, "Калужская область": 3009, "Костромская область": 2354, "Курская область": 2852,
    "Липецкая область": 3343, "Московская область": 5780, "Орловская область": 2529, "Рязанская область": 2607,
    "Смоленская область": 2571, "Тамбовская область": 2729, "Тверская область": 2657, "Тульская область": 2966,
    "Ярославская область": 3005, "Город Москва": 12500, "Республика Карелия": 3019, "Республика Коми": 3439,
    "Ненецкий автономный округ": 9913, "Архангельская область": 3515, "Вологодская область": 2537, "Калининградская область": 2749,
    "Ленинградская область": 3519, "Мурманская область": 4017, "Новгородская область": 2883, "Псковская область": 2304,
    "Город Санкт-Петербург": 6255, "Республика Адыгея": 2581, "Республика Калмыкия": 1644, "Республика Крым": 2108,
    "Краснодарский край": 4171, "Астраханская область": 2668, "Волгоградская область": 2657, "Ростовская область": 3502,
    "Город Севастополь": 2340, "Республика Дагестан": 2533, "Республика Ингушетия": 1021,
    "Кабардино-Балкарская Республика": 1984, "Карачаево-Черкесская Республика": 1439, "Республика Северная Осетия-Алания": 2273,
    "Чеченская Республика": 1926, "Ставропольский край": 2744, "Республика Башкортостан": 3457,
    "Республика Марий Эл": 1991, "Республика Мордовия": 1929, "Республика Татарстан": 4219, "Удмуртская Республика": 2613,
    "Чувашская Республика": 2041, "Пермский край": 3411, "Кировская область": 2365, "Нижегородская область": 3652,
    "Оренбургская область": 2775, "Пензенская область": 2460, "Самарская область": 3425, "Саратовская область": 2369,
    "Ульяновская область": 2348, "Курганская область": 2145, "Свердловская область": 4531,
    "Ханты-Мансийский автономный округ": 6770, "Ямало-Ненецкий автономный округ": 9790, "Тюменская область": 4253,
    "Челябинская область": 2984, "Республика Алтай": 2070, "Республика Бурятия": 2576, "Республика Тыва": 1562,
    "Республика Хакасия": 2547, "Алтайский край": 2330, "Забайкальский край": 2751, "Красноярский край": 3861,
    "Иркутская область": 3018, "Кемеровская область": 2727, "Новосибирская область": 3411, "Омская область": 2664,
    "Томская область": 2739, "Республика Саха (Якутия)": 5184, "Камчатский край": 4159, "Приморский край": 3505,
    "Хабаровский край": 4331, "Амурская область": 3153, "Магаданская область": 5268, "Сахалинская область": 6816,
    "Еврейская автономная область": 2528, "Чукотский автономный округ": 5853,
}


@dataclass(frozen=True)
class EventProfile:
    key: str
    provider: str
    profile_title: str
    music_label: str
    quarterly_fixed: int = 12_500
    max_duration_hours: int = 24
    extra_day_fee: int = 3_500
    special_hourly_limit: int = 8
    no_data_visitor_category: str = "unlimited"
    visits_point_label: str = "2.3"
    free_access_point_label: str = "2.2"
    rate_label_10k: str = "Применяемая ставка авторского вознаграждения (до 10 000 посещений)"
    rate_label_region: str = "Применяемая ставка авторского вознаграждения (согласно региону)"
    annual_fee_label: str = "Сумма авторского вознаграждения за год"


RAO_PROFILE = EventProfile(
    key="rao_events",
    provider="РАО",
    profile_title="Положение РАО для трансляторов мероприятий",
    music_label="произведений",
)


def _to_int(value: Any, default: int = 0) -> int:
    if value is None:
        return default
    if isinstance(value, bool):
        return int(value)
    s = str(value).strip().replace(" ", "")
    if not s:
        return default
    try:
        return int(float(s.replace(",", ".")))
    except Exception:
        return default


def _to_float(value: Any, default: float = 0.0) -> float:
    if value is None:
        return default
    s = str(value).strip().replace(" ", "").replace(",", ".")
    if not s:
        return default
    try:
        return float(s)
    except Exception:
        return default


def _parse_duration_hours(value: Any) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if not s:
        return 0.0
    # Поддержка форматов HH:MM:SS и MM:SS (если значение пришло строкой).
    if ":" in s:
        parts = [p.strip() for p in s.split(":")]
        try:
            nums = [float(p.replace(",", ".")) for p in parts]
        except Exception:
            nums = []
        if len(nums) == 3:
            h, m, sec = nums
            return max(0.0, h) + max(0.0, m) / 60.0 + max(0.0, sec) / 3600.0
        if len(nums) == 2:
            m, sec = nums
            return max(0.0, m) / 60.0 + max(0.0, sec) / 3600.0
    return _to_float(s, 0.0)


def _round_duration_hours(raw: Any) -> int:
    hours = _parse_duration_hours(raw)
    whole = int(math.floor(hours))
    frac = hours - whole
    rounded = whole + (1 if frac >= 0.5 else 0)
    return max(1, rounded)


def _duration_bucket(duration_hours: int) -> int:
    if duration_hours <= 4:
        return 4
    if duration_hours <= 8:
        return 8
    if duration_hours <= 12:
        return 12
    return 24


def _visitor_category(visitors: int, has_reliable_data: bool, no_data_category: str) -> str:
    if not has_reliable_data:
        return no_data_category
    if visitors <= 10_000:
        return "10000"
    if visitors <= 100_000:
        return "100000"
    if visitors <= 1_000_000:
        return "1000000"
    return "unlimited"


def _fmt_int(n: float | int) -> str:
    return f"{int(round(float(n))):,}".replace(",", " ")


def _build_text(
    profile: EventProfile,
    total: int,
    breakdown: list[dict[str, Any]],
    summary: dict[str, Any],
    contract_terms: list[dict[str, str]] | None = None,
) -> str:
    sep = "────────────────────────────────────────"
    lines = [
        sep,
        f"{profile.provider} — онлайн-трансляция мероприятия",
        sep,
        "",
        "ИСХОДНЫЕ ДАННЫЕ",
        f"• Режим расчёта: {summary.get('mode_label', '—')}.",
        f"• Доступ к трансляции: {summary.get('access_label', '—')}.",
        f"• Площадки: администрируемые — {summary.get('admin_count', 0)}, неадминистрируемые — {summary.get('external_count', 0)}.",
    ]
    share_label = summary.get("share_label")
    if share_label:
        lines.insert(7, f"• Доля {profile.music_label}: {share_label}.")
    if summary.get("duration_label"):
        lines.append(f"• Длительность: {summary['duration_label']}.")
    if summary.get("visitors_label"):
        lines.append(f"• Посещения: {summary['visitors_label']}.")
    lines.extend(
        [
            "",
            sep,
            "РАСЧЁТ",
            sep,
        ]
    )
    for i, item in enumerate(breakdown, start=1):
        formula = str(item.get("formula") or "—")
        lines.append(f"{i}. {item.get('label', 'Шаг расчёта')}")
        lines.append(f"   Формула: {formula}")
        if isinstance(item.get("value"), (int, float)):
            lines.append(f"   Результат: {_fmt_int(item['value'])} ₽")
    if contract_terms:
        lines.extend(
            [
                "",
                sep,
                "УСЛОВИЯ В ДОГОВОРЕ",
                sep,
            ]
        )
        for item in contract_terms:
            label = str(item.get("label") or "").strip()
            value = str(item.get("value") or "").strip()
            if not label:
                continue
            if value:
                lines.append(f"• {label}: {value}")
            else:
                lines.append(f"• {label}")
            note = str(item.get("note") or "").strip()
            if note:
                lines.append(f"  {note}")
    lines.extend(
        [
            "",
            sep,
            f"ИТОГО К ОПЛАТЕ: {_fmt_int(total)} ₽",
            sep,
        ]
    )
    return "\n".join(lines)


def calculate_event_fee(payload: dict[str, Any], profile: EventProfile = RAO_PROFILE) -> dict[str, Any]:
    user_type = str(payload.get("user_type") or "").strip()
    special_type = str(payload.get("special_calculation_type") or "").strip()
    has_reliable_data = bool(payload.get("has_reliable_data", True))
    visitors = max(0, _to_int(payload.get("visitors"), 0))
    duration = _to_float(payload.get("duration_hours"), 0.0)
    duration_special = _to_float(payload.get("duration_hours_special"), 0.0)
    region = str(payload.get("region") or "").strip()
    is_free_access = bool(payload.get("is_free_access", False))
    music_share = str(payload.get("music_share") or "100").strip()
    admin_platforms_count = max(0, _to_int(payload.get("admin_platforms_count"), 0))
    external_platforms_count = max(0, _to_int(payload.get("external_platforms_count"), 0))
    broadcasts_count = max(1, _to_int(payload.get("broadcasts_count"), 1))
    quarterly_contract_year_confirmed = bool(payload.get("quarterly_contract_year_confirmed", True))
    quarterly_prepay_4q_confirmed = bool(payload.get("quarterly_prepay_4q_confirmed", True))

    if user_type not in {"standard", "special"}:
        raise ValueError("Некорректный тип пользователя.")
    if user_type == "special" and special_type not in {"quarterly", "hourly"}:
        raise ValueError("Для специальной категории выберите вариант расчёта (3.2.1 или 3.2.2).")

    breakdown: list[dict[str, Any]] = []
    contract_terms: list[dict[str, str]] = []
    total = 0.0
    is_special_quarterly = user_type == "special" and special_type == "quarterly"
    platform_multiplier = 0

    def _base_amount_for_category(visitor_key: str, rounded_hours: int) -> tuple[float, int]:
        if rounded_hours <= profile.max_duration_hours:
            bucket = _duration_bucket(rounded_hours)
            rate_hour = EVENT_RATES_TABLE[visitor_key][bucket]
            return float(rate_hour * rounded_hours), int(rate_hour)
        first24_rate = EVENT_RATES_TABLE[visitor_key][24]
        first24 = first24_rate * profile.max_duration_hours
        full_extra_days = max(0, (rounded_hours - profile.max_duration_hours) // profile.max_duration_hours)
        extra = full_extra_days * profile.extra_day_fee
        return float(first24 + extra), int(first24_rate)

    if user_type == "special" and special_type == "quarterly":
        if not quarterly_contract_year_confirmed:
            raise ValueError("Режим 3.2.1 требует договор на срок не менее одного года.")
        if not quarterly_prepay_4q_confirmed:
            raise ValueError("Режим 3.2.1 требует единовременную оплату за 4 календарных квартала.")
        total = profile.quarterly_fixed * 4
        breakdown.append(
            {
                "label": "Фиксированная ставка",
                "formula": f"{_fmt_int(profile.quarterly_fixed)} ₽ × 4 квартала",
                "value": total,
            }
        )
        admin_part_q = 1 if admin_platforms_count > 0 else 0
        platforms_q = admin_part_q + external_platforms_count
        if platforms_q <= 0:
            platforms_q = 1
        platform_multiplier = platforms_q
        annual_with_platforms = (profile.quarterly_fixed * 4) * platforms_q
        total = annual_with_platforms
        breakdown.append(
            {
                "label": "Учёт площадок",
                "formula": f"{_fmt_int(profile.quarterly_fixed * 4)} ₽ × {platforms_q} = {_fmt_int(annual_with_platforms)} ₽ за год",
                "value": total,
            }
        )
        contract_terms.append(
            {
                "label": profile.annual_fee_label,
                "value": f"{_fmt_int(annual_with_platforms)} ₽",
            }
        )
    elif user_type == "special" and special_type == "hourly":
        if not region or region not in EVENT_REGION_RATES:
            raise ValueError("Для расчёта по 3.2.2 необходимо выбрать регион.")
        if duration_special <= 0:
            raise ValueError("Укажите длительность трансляции больше нуля.")

        rounded = _round_duration_hours(duration_special)
        effective_hours = min(rounded, profile.special_hourly_limit)
        hourly_rate = EVENT_REGION_RATES[region]
        total = hourly_rate * effective_hours
        breakdown.append(
            {
                "label": "Базовая сумма (3.2.2)",
                "formula": f"{_fmt_int(hourly_rate)} ₽ × {effective_hours} ч.",
                "value": total,
            }
        )
        if rounded > profile.special_hourly_limit:
            breakdown.append(
                {
                    "label": "Ограничение длительности (п.3.3)",
                    "formula": f"Учитываются только первые {profile.special_hourly_limit} ч. из {rounded} ч.",
                    "value": total,
                }
            )
        base_without_share = total
    else:
        if duration <= 0:
            raise ValueError("Укажите длительность трансляции больше нуля.")
        if has_reliable_data and payload.get("visitors") in ("", None):
            raise ValueError("Укажите количество посещений или отметьте отсутствие достоверных данных.")

        rounded = _round_duration_hours(duration)
        visitor_category = _visitor_category(visitors, has_reliable_data, profile.no_data_visitor_category)
        duration_category = _duration_bucket(rounded)
        hourly_rate = EVENT_RATES_TABLE[visitor_category][duration_category]

        if rounded <= profile.max_duration_hours:
            total = hourly_rate * rounded
            breakdown.append(
                {
                    "label": "Базовая сумма",
                    "formula": f"{_fmt_int(hourly_rate)} ₽ × {rounded} ч.",
                    "value": total,
                }
            )
        else:
            first24_rate = EVENT_RATES_TABLE[visitor_category][24]
            first24 = first24_rate * profile.max_duration_hours
            full_extra_days = max(0, (rounded - profile.max_duration_hours) // profile.max_duration_hours)
            extra = full_extra_days * profile.extra_day_fee
            total = first24 + extra
            breakdown.append(
                {
                    "label": "Первые 24 часа",
                    "formula": f"{_fmt_int(first24_rate)} ₽ × 24 ч.",
                    "value": first24,
                }
            )
            if full_extra_days > 0:
                breakdown.append(
                    {
                        "label": "Полные дополнительные сутки (п.2.2)",
                        "formula": f"{_fmt_int(profile.extra_day_fee)} ₽ × {full_extra_days} сут.",
                        "value": extra,
                    }
                )
            tail_hours = max(0, rounded - profile.max_duration_hours - full_extra_days * profile.max_duration_hours)
            if tail_hours > 0:
                breakdown.append(
                    {
                        "label": "Неполные сутки после первых 24 часов",
                        "formula": (
                            f"После первых 24 ч. осталось {tail_hours} ч.; "
                            "доплата 3 500 ₽ применяется только за полные сутки."
                        ),
                        "value": total,
                    }
                )
        if not has_reliable_data:
            breakdown.append(
                {
                    "label": f"Количество посещений (п.{profile.visits_point_label})",
                    "formula": "Поскольку достоверные данные о посещениях отсутствуют, расчёт производится исходя из максимальной ставки.",
                    "value": total,
                }
            )
        std_base_10k, std_rate_10k = _base_amount_for_category("10000", rounded)
        std_base_max, _std_rate_max = _base_amount_for_category("unlimited", rounded)

    if not is_special_quarterly:
        coef = EVENT_MUSIC_SHARE_COEFFICIENTS.get(music_share, 1.0)
        if music_share == "unknown":
            breakdown.append(
                {
                    "label": f"Коэффициент доли {profile.music_label} (п.4.3)",
                    "formula": "Доля не подтверждена; понижающий коэффициент не применён (коэффициент 1,0).",
                    "value": total,
                }
            )
        elif coef != 1.0:
            before = total
            total = total * coef
            breakdown.append(
                {
                    "label": f"Коэффициент доли {profile.music_label} (п.4.3)",
                    "formula": f"{_fmt_int(before)} ₽ × {coef}",
                    "value": total,
                }
            )

        if user_type == "standard" and is_free_access:
            before = total
            total = total * 0.5
            breakdown.append(
                {
                    "label": f"Коэффициент бесплатного доступа (п.{profile.free_access_point_label})",
                    "formula": f"{_fmt_int(before)} ₽ × 0,5",
                    "value": total,
                }
            )
        elif user_type == "special" and special_type == "hourly" and is_free_access:
            breakdown.append(
                {
                    "label": "Коэффициент бесплатного доступа",
                    "formula": "Не применён: для специального режима 3.2.2 пункт 2.2 не используется.",
                    "value": total,
                }
            )

        if user_type != "standard" and broadcasts_count > 1:
            before = total
            total = total * broadcasts_count
            breakdown.append(
                {
                    "label": "Количество трансляций мероприятия",
                    "formula": f"{_fmt_int(before)} ₽ × {broadcasts_count} трансляций.",
                    "value": total,
                }
            )

    if not is_special_quarterly:
        admin_part = 1 if admin_platforms_count > 0 else 0
        platform_multiplier = admin_part + external_platforms_count
        if platform_multiplier <= 0:
            raise ValueError("Укажите хотя бы одну площадку трансляции (администрируемую или неадминистрируемую).")
        if platform_multiplier > 1:
            before = total
            total = total * platform_multiplier
            breakdown.append(
                {
                    "label": "Количество площадок (п.4.2)",
                    "formula": (
                        f"{_fmt_int(before)} ₽ × {platform_multiplier} "
                        f"(администрируемые: {admin_platforms_count}, неадминистрируемые: {external_platforms_count})"
                    ),
                    "value": total,
                }
            )
        else:
            breakdown.append(
                {
                    "label": "Учет площадок (п.4.2)",
                    "formula": "Администрируемые площадки учитываются как одна единица.",
                    "value": total,
                }
            )

    if user_type == "standard":
        free_coef = 0.5 if is_free_access else 1.0
        extra_coef = float(platform_multiplier if platform_multiplier > 0 else 1)
        sum_10k = std_base_10k * free_coef * extra_coef
        sum_10k_60 = sum_10k * 0.6
        sum_max = std_base_max * free_coef * extra_coef
        contract_terms.extend(
            [
                {
                    "label": profile.rate_label_10k,
                    "value": f"{_fmt_int(sum_10k)} ₽",
                    "note": "Сумма ставки с учётом длительности, платного/бесплатного доступа и количества площадок.",
                },
                {
                    "label": "Сумма в размере 60% от суммы ставки, учитывающей не более 10 000 посещений",
                    "value": f"{_fmt_int(sum_10k_60)} ₽",
                },
                {
                    "label": "Сумма исходя из максимального количества посещений (свыше 1 000 000)",
                    "value": f"{_fmt_int(sum_max)} ₽",
                },
            ]
        )
    elif user_type == "special" and special_type == "hourly":
        base_without_share_total = (base_without_share if "base_without_share" in locals() else total)
        base_without_share_total *= float(platform_multiplier if platform_multiplier > 0 else 1)
        if broadcasts_count > 1:
            base_without_share_total *= float(broadcasts_count)
        region_rate = EVENT_REGION_RATES.get(region, 0)
        contract_terms.extend(
            [
                {
                    "label": profile.rate_label_region,
                    "value": f"{_fmt_int(region_rate)} ₽/ч",
                },
                {
                    "label": "Сумма без понижающего коэффициента",
                    "value": f"{_fmt_int(base_without_share_total)} ₽",
                },
                {
                    "label": "60% от суммы с учётом количества часов и площадок",
                    "value": f"{_fmt_int(base_without_share_total * 0.6)} ₽",
                },
            ]
        )

    final_total = int(round(total))
    if user_type == "special" and special_type == "quarterly":
        mode_label = "Специальная категория 3.2.1 (12 500 ₽ за квартал)"
    elif user_type == "special" and special_type == "hourly":
        mode_label = "Специальная категория 3.2.2 (региональная почасовая ставка)"
    else:
        mode_label = "Обычный режим (посещения и длительность)"

    if user_type == "special" and special_type == "quarterly":
        access_label = "Не учитывается в спецкатегории 3.2.1"
    elif user_type == "special" and special_type == "hourly":
        access_label = "Не применяется к режиму 3.2.2"
    else:
        access_label = "Бесплатный" if is_free_access else "Платный"

    visitors_label = None
    if user_type == "standard":
        visitors_label = (
            f"{_fmt_int(visitors)} (достоверные данные)"
            if has_reliable_data
            else "Нет достоверных данных, поэтому при расчёте применяется максимальная ставка"
        )

    duration_label = None
    if user_type == "special" and special_type == "hourly":
        rounded_special = _round_duration_hours(duration_special)
        duration_label = f"{duration_special:g} ч. (округлено: {rounded_special} ч., учитывается не более 8 ч.)"
    elif user_type == "standard":
        rounded_std = _round_duration_hours(duration)
        duration_label = f"{duration:g} ч. (округлено: {rounded_std} ч.)"

    summary = {
        "mode_label": mode_label,
        "access_label": access_label,
        "share_label": (
            None
            if is_special_quarterly
            else (
                "Доля не подтверждена (коэффициент не применён)"
                if music_share == "unknown"
                else (
                    "80% и более"
                    if music_share == "100"
                    else f"до {music_share}%"
                )
            )
        ),
        "admin_count": admin_platforms_count,
        "external_count": external_platforms_count,
        "duration_label": duration_label,
        "visitors_label": visitors_label,
    }
    text = _build_text(profile, final_total, breakdown, summary, contract_terms=contract_terms)
    return {
        "ok": True,
        "calculator": profile.key,
        "provider": profile.provider,
        "total": final_total,
        "breakdown": breakdown,
        "contract_terms": contract_terms,
        "text": text,
        "normalized": {
            "user_type": user_type,
            "special_calculation_type": special_type or None,
            "visitors": visitors if has_reliable_data else None,
            "duration_hours": duration if duration > 0 else None,
            "duration_hours_special": duration_special if duration_special > 0 else None,
            "region": region or None,
            "music_share": music_share,
            "is_free_access": is_free_access,
            "admin_platforms_count": admin_platforms_count,
            "external_platforms_count": external_platforms_count,
            "broadcasts_count": broadcasts_count,
            "quarterly_contract_year_confirmed": quarterly_contract_year_confirmed,
            "quarterly_prepay_4q_confirmed": quarterly_prepay_4q_confirmed,
        },
    }


__all__ = [
    "CALCULATOR_ID",
    "PROVIDER",
    "RAO_PROFILE",
    "EVENT_REGION_RATES",
    "EVENT_MUSIC_SHARE_COEFFICIENTS",
    "calculate_event_fee",
]
