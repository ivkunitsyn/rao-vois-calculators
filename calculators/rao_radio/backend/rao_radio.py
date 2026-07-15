# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import io
import math
import html as html_lib
from contextlib import redirect_stdout
from decimal import Decimal
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple
import re

import pandas as pd

from calculators.rao_tv.backend import rao as base
from calculators.rao_radio.result.model import (
    ContractTerms,
    MinimumLicense,
    RadioReportModel,
    RateChannel,
    RateLicense,
    RateTopic,
    ReportRow,
)
from calculators.rao_radio.result.formatting import (
    format_financial_base_type,
    format_internet_component_label,
    format_internet_income_threshold_label,
)
from calculators.rao_radio.result.renderer_html import render_radio_report_html


# --------------------------- constants ---------------------------

GENERAL_RAO_RATES: Dict[str, float] = {
    "I": 1.6,
    "II": 2.0,
    "III": 2.3,
    "IV": 2.7,
    "V": 3.0,
    "UNKNOWN": 2.3,
}

GENERAL_VOIS_RATES: Dict[str, float] = {
    "I": 1.1,
    "II": 1.5,
    "III": 1.8,
    "IV": 2.2,
    "V": 2.5,
    "UNKNOWN": 1.8,
}

RAR_RAO_RATES: Dict[str, float] = {
    "I": 1.6,
    "II": 1.9,
    "III": 2.2,
    "IV": 2.5,
    "V": 2.7,
    "UNKNOWN": 2.2,
}

# Значения из пользовательской спецификации.
GENERAL_RAO_MINIMUMS: List[Tuple[Tuple[float, float], List[float]]] = [
    ((1.6, 1.8), [2500, 5500, 10000, 18500, 25000, 32000, 52500, 105000]),
    ((1.9, 2.1), [3000, 6200, 11500, 19000, 26000, 33000, 66000, 120000]),
    ((2.2, 2.4), [3500, 6600, 13500, 20000, 27500, 36000, 74000, 135000]),
    ((2.5, 2.7), [3900, 7500, 16000, 22000, 29000, 40000, 85000, 150000]),
    ((2.8, 3.0), [4500, 8500, 18000, 24500, 31500, 45000, 95000, 170000]),
]

RAR_RAO_MINIMUMS: List[Tuple[Tuple[float, float], List[float]]] = [
    ((1.6, 1.7), [2125, 4675, 8245, 15725, 21250, 27200, 44625, 89250]),
    ((1.8, 2.0), [2550, 5270, 9775, 16150, 22100, 28050, 56100, 102000]),
    ((2.1, 2.3), [2975, 5610, 11475, 17000, 23375, 30600, 62900, 114750]),
    ((2.4, 2.5), [3315, 6375, 13600, 18700, 24650, 34000, 72250, 127500]),
    ((2.6, 2.7), [3825, 7225, 14450, 20825, 26775, 38250, 80750, 144500]),
]

GENERAL_VOIS_MINIMUMS: List[Tuple[Tuple[float, float], List[float]]] = [
    ((1.1, 1.3), [2300, 5000, 9000, 13500, 21000, 27000, 51000, 93000]),
    ((1.4, 1.6), [2700, 5500, 10000, 14500, 22000, 30000, 60000, 100000]),
    ((1.7, 1.9), [3200, 6500, 10500, 16500, 23000, 33000, 70000, 115000]),
    ((2.0, 2.2), [3500, 7200, 11700, 18000, 25000, 39000, 80000, 130000]),
    ((2.3, 2.5), [4000, 8300, 13000, 20500, 26500, 44000, 90000, 145000]),
]

POP_BUCKETS: List[Tuple[int, Optional[int]]] = [
    (0, 49_990),
    (50_000, 149_990),
    (150_000, 299_990),
    (300_000, 599_990),
    (600_000, 999_990),
    (1_000_000, 1_999_990),
    (2_000_000, 3_999_990),
    (4_000_000, None),
]

MAX_REASONABLE_RUSSIA_POPULATION = 146_500_000


# --------------------------- models ---------------------------

@dataclass
class RadioTopic:
    name: str
    category: str
    share_percent: Optional[float] = None


@dataclass
class RadioChannel:
    name: str
    weekly_hours: float
    topics: List[RadioTopic] = field(default_factory=list)
    actual_share_percent: Optional[float] = None
    founder_type: str = "user"  # user | third_party
    simultaneous_internet_broadcast: bool = False
    internet_admin_by_user: bool = True
    site_app_count: int = 0
    notes: List[str] = field(default_factory=list)


@dataclass
class RadioLicense:
    license_id: str
    org_name: str
    media_raw: str
    population: Optional[int]
    license_date: str = ""
    service_start_date: str = ""
    channels: List[RadioChannel] = field(default_factory=list)
    population_notes: List[str] = field(default_factory=list)
    rkn_url: str = ""

    def weekly_hours(self) -> float:
        if not self.channels:
            return 168.0
        return float(sum(max(0.0, float(ch.weekly_hours or 0.0)) for ch in self.channels))


@dataclass
class UserFlags:
    status_code: str
    new_license_ids: Set[str]
    is_rar_member: bool
    is_state_budget_institution: bool
    is_internet_only: bool
    has_documented_income: bool
    is_100_state_capital: bool
    is_new_user_initiated_contract: bool
    report_period_number_from_contract_start: int
    is_new_user_and_other_use_contract: bool
    is_package_contract: bool
    assoc_member: bool
    has_legal_cases: bool
    signed_within_30_days: bool
    has_current_agreement_breach: bool
    report_period_number_from_special_minimum_start: int
    previous_year_income_or_expense: float
    can_use_fixed_fee: bool
    fixed_fee_basis_confirmed: bool
    market_index: float
    inflation_index: float
    simultaneous_internet_broadcast: bool
    internet_admin_by_user: bool
    site_app_count: int
    has_third_party_channels: bool


# --------------------------- topic category ---------------------------

def _load_topic_map(vars_xlsx: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(vars_xlsx, sheet_name="Тематики по категориям")
        if df.empty:
            return pd.DataFrame(columns=["topic_norm", "category"])
        src_col = base.TOPIC_MAP_COL_TOPIC
        cat_col = base.TOPIC_MAP_COL_CAT
        if src_col not in df.columns or cat_col not in df.columns:
            return pd.DataFrame(columns=["topic_norm", "category"])
        out = df[[src_col, cat_col]].dropna().copy()
        out["topic_norm"] = out[src_col].astype(str).map(base.normalize_topic)
        out["category"] = out[cat_col].astype(str).str.strip().str.upper()
        out = out[(out["topic_norm"] != "") & (out["category"] != "")]
    except Exception:
        out = pd.DataFrame(columns=["topic_norm", "category"])

    # Добавляем расширенный словарь из ТВ-модуля.
    for cat, topics in base.EXTRA_TOPICS_BY_CATEGORY.items():
        for t in topics:
            out.loc[len(out)] = {"topic_norm": base.normalize_topic(t), "category": str(cat).strip().upper()}
    return out


def detect_topic_category(topic_name: str, topic_map: pd.DataFrame) -> str:
    t_norm = base.normalize_topic(topic_name or "")
    if not t_norm:
        return "UNKNOWN"

    if topic_map is not None and not topic_map.empty:
        exact = topic_map[topic_map["topic_norm"] == t_norm]
        if not exact.empty:
            cat = str(exact.iloc[0]["category"]).upper()
            return cat if cat in {"I", "II", "III", "IV", "V"} else "UNKNOWN"

    tl = t_norm
    def hit(*keys: str) -> bool:
        return any(k in tl for k in keys)

    # Для РВ в составных формулировках без точного совпадения берём элемент
    # с наибольшей долей музыкального вещания. Музыкальные тематики имеют приоритет.
    if hit("музык", "музыаль", "песн", "концерт", "музыка"):
        return "V"

    if topic_map is not None and not topic_map.empty:
        cands = topic_map[topic_map["topic_norm"].apply(lambda x: bool(x) and (x in t_norm or t_norm in x))]
        if not cands.empty:
            cands = cands.copy()
            cands["_len"] = cands["topic_norm"].str.len()
            cands = cands.sort_values("_len", ascending=False)
            cat = str(cands.iloc[0]["category"]).upper()
            return cat if cat in {"I", "II", "III", "IV", "V"} else "UNKNOWN"

    # Упрощённые эвристики.
    if hit("реклам"):
        return "V"
    if hit("развлек", "шоу", "юмор", "конкурс", "игров"):
        return "IV"
    if hit("культур", "художествен", "документ", "литератур", "искусств", "просвет"):
        return "III"
    if hit("дет", "образоват", "науч", "познав", "спорт", "оздоров", "зож", "туризм"):
        return "II"
    if hit("информац", "новост", "полит", "эконом", "прав", "публицист", "религи", "социаль"):
        return "I"
    return "UNKNOWN"


# --------------------------- loaders ---------------------------

def _is_radio_license(media_raw: str, licensed_activity: str = "") -> bool:
    a = str(licensed_activity or "").lower()
    if "радио" in a or "радиоканал" in a or "радиовещ" in a:
        return True
    s = str(media_raw or "").lower()
    return "радио" in s or "радиоканал" in s or "радиовещ" in s


def _is_tv_activity(licensed_activity: str = "") -> bool:
    a = str(licensed_activity or "").lower()
    return ("теле" in a) or ("телеканал" in a) or ("телевещ" in a)


def _is_explicit_tv_license(media_raw: str, licensed_activity: str = "") -> bool:
    a = str(licensed_activity or "").lower()
    a_has_radio = ("радио" in a) or ("радиоканал" in a) or ("радиовещ" in a)
    a_has_tv = ("теле" in a) or ("телеканал" in a) or ("телевещ" in a)
    if a_has_radio:
        return False
    if a_has_tv:
        return True
    s = str(media_raw or "").lower()
    s_has_radio = ("радио" in s) or ("радиоканал" in s) or ("радиовещ" in s)
    s_has_tv = ("теле" in s) or ("телеканал" in s) or ("телевещ" in s)
    return s_has_tv and (not s_has_radio)


def _cap_population_for_radio(population: Optional[int], notes: List[str]) -> Optional[int]:
    if population is None:
        return None
    p = int(population)
    if p > MAX_REASONABLE_RUSSIA_POPULATION:
        notes.append(
            f"Численность населения по лицензии превышала численность РФ ({format_count(p)}); "
            f"для расчёта использовано {format_count(MAX_REASONABLE_RUSSIA_POPULATION)}."
        )
        return MAX_REASONABLE_RUSSIA_POPULATION
    return p


def load_radio_licenses_by_inn(
    rkn_xlsx: Path,
    inn: str,
    vars_xlsx: Path,
    runtime_population_normalization: bool = True,
) -> Tuple[List[RadioLicense], List[str]]:
    notes: List[str] = []
    topic_map = _load_topic_map(vars_xlsx)

    src_licenses, src_notes = base.load_licenses_by_inn(
        rkn_xlsx=rkn_xlsx,
        inn=inn,
        vars_xlsx=vars_xlsx,
        filter_radio_channels=False,
        runtime_population_normalization=runtime_population_normalization,
    )
    notes.extend(src_notes or [])

    out: List[RadioLicense] = []
    for lic in src_licenses:
        licensed_activity = getattr(lic, "licensed_activity", "")
        lic_is_radio = _is_radio_license(lic.media_raw, licensed_activity)
        explicit_tv = _is_explicit_tv_license(lic.media_raw, licensed_activity)
        radio_named_channels = [ch for ch in lic.channels if base.is_radio_channel_name(ch.name)]
        has_radio_flag = bool(radio_named_channels)
        media_l = str(lic.media_raw or "").lower()
        media_is_radio = ("радио" in media_l) or ("радиоканал" in media_l) or ("радиовещ" in media_l)
        media_is_tv = ("теле" in media_l) or ("телеканал" in media_l) or ("телевещ" in media_l)
        activity_is_radio = _is_radio_license("", licensed_activity)
        activity_is_tv = _is_tv_activity(licensed_activity)

        # Синхронизируем правила с /api/licenses (radio_only), чтобы
        # выбранная на шаге лицензий запись не "пропадала" в финальном расчёте.
        if activity_is_tv and not activity_is_radio:
            continue
        if (
            not activity_is_radio
            and (media_is_tv and not media_is_radio and not has_radio_flag)
        ):
            continue
        if explicit_tv and not has_radio_flag and not activity_is_radio:
            continue

        if radio_named_channels:
            selected_channels = radio_named_channels
        elif lic_is_radio or activity_is_radio:
            selected_channels = list(lic.channels)
        else:
            # Для универсальных/неполных выгрузок без явной маркировки
            # радиовещания берём каналы лицензии, если она не распознана как ТВ.
            selected_channels = list(lic.channels)

        if not selected_channels:
            continue

        population_notes = list(lic.population_notes or [])
        population = _cap_population_for_radio(lic.population_total, population_notes)

        radio_channels: List[RadioChannel] = []
        for ch in selected_channels:
            topics: List[RadioTopic] = []
            for t in ch.topics:
                cat = detect_topic_category(t.topic_raw, topic_map)
                topics.append(RadioTopic(name=t.topic_raw, category=cat, share_percent=t.share_pct))
            radio_channels.append(
                RadioChannel(
                    name=ch.name,
                    weekly_hours=float(ch.hours_week or 168.0),
                    topics=topics,
                    notes=list(ch.hours_notes or []),
                )
            )

        out.append(
            RadioLicense(
                license_id=lic.license_id,
                org_name=lic.org_name,
                media_raw=lic.media_raw,
                population=population,
                license_date=str(getattr(lic, "license_date", "") or ""),
                service_start_date=str(getattr(lic, "service_start_date", "") or ""),
                channels=radio_channels,
                population_notes=population_notes,
                rkn_url=lic.rkn_url,
            )
        )

    if not out:
        notes.append("По ИНН не найдено действующих лицензий на осуществление радиовещания.")
    return out, notes


# --------------------------- calculators ---------------------------

def round_rate(x: float) -> float:
    return round(float(x) + 1e-9, 1)


def use_rar_rates(u: UserFlags) -> bool:
    return bool(u.is_rar_member and (not u.is_state_budget_institution) and (not u.is_internet_only))


def _is_vois_society(society: str) -> bool:
    return str(society or "").strip().upper() == "ВОИС"


def _rate_table(use_rar: bool, is_vois: bool) -> Dict[str, float]:
    if is_vois:
        return GENERAL_VOIS_RATES
    return RAR_RAO_RATES if use_rar else GENERAL_RAO_RATES


def rate_by_share(share: float, use_rar: bool, is_vois: bool) -> float:
    x = float(share)
    if is_vois:
        if x <= 19.99:
            return 1.1
        if x <= 39.99:
            return 1.5
        if x <= 59.99:
            return 1.8
        if x <= 79.99:
            return 2.2
        return 2.5
    if use_rar:
        if x <= 19.99:
            return 1.6
        if x <= 39.99:
            return 1.9
        if x <= 59.99:
            return 2.2
        if x <= 79.99:
            return 2.5
        return 2.7
    if x <= 19.99:
        return 1.6
    if x <= 39.99:
        return 2.0
    if x <= 59.99:
        return 2.3
    if x <= 79.99:
        return 2.7
    return 3.0


def calculate_channel_rate(ch: RadioChannel, use_rar: bool, is_vois: bool) -> float:
    table = _rate_table(use_rar, is_vois)
    if ch.actual_share_percent is not None:
        return round_rate(rate_by_share(ch.actual_share_percent, use_rar, is_vois))
    if not ch.topics:
        return round_rate(table["UNKNOWN"])

    with_share = [t for t in ch.topics if t.share_percent is not None]
    if with_share:
        dominant = next((t for t in with_share if float(t.share_percent or 0.0) > 50.0), None)
        if dominant is not None:
            return round_rate(table.get(dominant.category, table["UNKNOWN"]))
        s = sum(max(0.0, float(t.share_percent or 0.0)) for t in with_share)
        if s > 0:
            w = sum(float(t.share_percent or 0.0) * table.get(t.category, table["UNKNOWN"]) for t in with_share) / s
            return round_rate(w)

    avg = sum(table.get(t.category, table["UNKNOWN"]) for t in ch.topics) / max(1, len(ch.topics))
    return round_rate(avg)


def calculate_license_rate(lic: RadioLicense, use_rar: bool, is_vois: bool) -> float:
    if not lic.channels:
        return round_rate(_rate_table(use_rar, is_vois)["UNKNOWN"])
    den = sum(max(0.0, float(ch.weekly_hours or 0.0)) for ch in lic.channels)
    if den <= 0:
        return round_rate(_rate_table(use_rar, is_vois)["UNKNOWN"])
    num = 0.0
    for ch in lic.channels:
        num += calculate_channel_rate(ch, use_rar, is_vois) * float(ch.weekly_hours or 0.0)
    return round_rate(num / den)


def calculate_contract_rate(licenses: List[RadioLicense], use_rar: bool, is_vois: bool) -> float:
    num = 0.0
    den = 0.0
    for lic in licenses:
        if lic.population is None:
            continue
        weight = float(lic.weekly_hours()) * float(lic.population)
        if weight <= 0:
            continue
        num += calculate_license_rate(lic, use_rar, is_vois) * weight
        den += weight
    if den <= 0:
        return round_rate(_rate_table(use_rar, is_vois)["UNKNOWN"])
    return round_rate(num / den)


def get_population_bucket_index(population: int) -> int:
    p = int(max(0, population))
    for i, (lo, hi) in enumerate(POP_BUCKETS):
        if p < lo:
            continue
        if hi is None or p <= hi:
            return i
    return len(POP_BUCKETS) - 1


def _pick_min_row(rate: float, use_rar: bool, is_vois: bool) -> List[float]:
    if is_vois:
        table = GENERAL_VOIS_MINIMUMS
    else:
        table = RAR_RAO_MINIMUMS if use_rar else GENERAL_RAO_MINIMUMS
    r = float(rate)
    for (lo, hi), vals in table:
        if lo <= r <= hi:
            return vals
    # fallback ближайший диапазон
    if r < table[0][0][0]:
        return table[0][1]
    return table[-1][1]


def _population_bucket_label(population: Optional[int]) -> str:
    if population is None:
        return "без населения"
    idx = get_population_bucket_index(int(population))
    lo, hi = POP_BUCKETS[idx]
    if hi is None:
        return f"от {format_count(lo)}"
    return f"{format_count(lo)}-{format_count(hi)}"


def get_hours_coefficient(weekly_hours: float) -> float:
    h = float(weekly_hours or 0.0)
    if h > 126:
        return 1.0
    if h >= 84:
        return 0.8
    if h >= 28:
        return 0.6
    return 0.4


def get_internet_addon_per_site(previous_year_income: float) -> float:
    x = float(previous_year_income or 0.0)
    if x <= 3_000_000:
        return 6_250.0
    if x <= 50_000_000:
        return 12_500.0
    return 25_000.0


def apply_own_internet_uplift(minimum: float, site_app_count: int, previous_year_income: float) -> float:
    by_coef = float(minimum) * 1.15
    addon = get_internet_addon_per_site(previous_year_income) * max(0, int(site_app_count))
    by_min_addon = float(minimum) + addon
    return max(by_coef, by_min_addon)


def apply_third_party_internet_uplift(minimum: float) -> float:
    return float(minimum) * 1.15


def _channel_minimum_inside_license(ch: RadioChannel, lic: RadioLicense, use_rar: bool, is_vois: bool) -> float:
    if lic.population is None:
        return 0.0
    ch_rate = calculate_channel_rate(ch, use_rar, is_vois)
    row = _pick_min_row(ch_rate, use_rar, is_vois)
    pop_idx = get_population_bucket_index(int(lic.population))
    base_min = float(row[pop_idx])
    base_min *= get_hours_coefficient(float(ch.weekly_hours or 0.0))
    return base_min


def calculate_base_minimum_by_license(lic: RadioLicense, use_rar: bool, is_vois: bool) -> float:
    if lic.population is None:
        return 0.0
    lic_rate = calculate_license_rate(lic, use_rar, is_vois)
    row = _pick_min_row(lic_rate, use_rar, is_vois)
    pop_idx = get_population_bucket_index(int(lic.population))
    return float(row[pop_idx])


def calculate_license_minimum(
    lic: RadioLicense,
    user: UserFlags,
    use_rar: bool,
    is_vois: bool,
) -> float:
    has_own = any(ch.founder_type == "user" for ch in lic.channels)
    has_third = any(ch.founder_type == "third_party" for ch in lic.channels)

    if has_own and has_third:
        total = 0.0
        for ch in lic.channels:
            m = _channel_minimum_inside_license(ch, lic, use_rar, is_vois)
            if ch.simultaneous_internet_broadcast and ch.internet_admin_by_user:
                if ch.founder_type == "user":
                    m = apply_own_internet_uplift(m, ch.site_app_count, user.previous_year_income_or_expense)
                else:
                    m = apply_third_party_internet_uplift(m)
            total += m
        return total

    m = calculate_base_minimum_by_license(lic, use_rar, is_vois)
    m *= get_hours_coefficient(lic.weekly_hours())

    has_own_internet = any(
        (ch.founder_type == "user" and ch.simultaneous_internet_broadcast and ch.internet_admin_by_user)
        for ch in lic.channels
    )
    has_third_internet = any(
        (ch.founder_type == "third_party" and ch.simultaneous_internet_broadcast and ch.internet_admin_by_user)
        for ch in lic.channels
    )
    if has_own_internet:
        own_site_count = max([max(0, int(ch.site_app_count)) for ch in lic.channels if ch.founder_type == "user"] or [0])
        m = apply_own_internet_uplift(m, own_site_count, user.previous_year_income_or_expense)
    if has_third_internet:
        m = apply_third_party_internet_uplift(m)
    return m


def calculate_license_minimum_without_internet(
    lic: RadioLicense,
    user: UserFlags,
    use_rar: bool,
    is_vois: bool,
) -> float:
    offline_channels = [
        RadioChannel(
            name=ch.name,
            weekly_hours=ch.weekly_hours,
            topics=list(ch.topics or []),
            actual_share_percent=ch.actual_share_percent,
            founder_type=ch.founder_type,
            simultaneous_internet_broadcast=False,
            internet_admin_by_user=ch.internet_admin_by_user,
            site_app_count=0,
            notes=list(ch.notes or []),
        )
        for ch in (lic.channels or [])
    ]
    offline_license = RadioLicense(
        license_id=lic.license_id,
        org_name=lic.org_name,
        media_raw=lic.media_raw,
        population=lic.population,
        license_date=lic.license_date,
        service_start_date=lic.service_start_date,
        channels=offline_channels,
        population_notes=list(lic.population_notes or []),
        rkn_url=lic.rkn_url,
    )
    return calculate_license_minimum(offline_license, user, use_rar, is_vois)


def reductions_allowed(user: UserFlags) -> bool:
    if user.has_legal_cases:
        return False
    if not user.signed_within_30_days:
        return False
    if user.has_current_agreement_breach:
        return False
    return True


def get_general_reduction_coefficient(user: UserFlags, use_rar: bool) -> float:
    cands: List[float] = []
    if user.is_new_user_and_other_use_contract:
        cands.append(0.95)
    if user.is_package_contract:
        cands.append(0.9)
    if user.assoc_member and not use_rar:
        cands.append(0.85)
    if not cands:
        return 1.0
    return min(cands)


def get_new_contract_coefficient(user: UserFlags) -> float:
    if not user.is_new_user_initiated_contract:
        return 1.0
    q = int(user.report_period_number_from_contract_start or 0)
    if 1 <= q <= 4:
        return 0.75
    if 5 <= q <= 8:
        return 0.88
    return 1.0


def get_new_contract_coefficient_by_period(period_number: int) -> float:
    q = int(period_number or 0)
    if 1 <= q <= 4:
        return 0.75
    if 5 <= q <= 8:
        return 0.88
    return 1.0


def _parse_registry_date(raw: str) -> Optional[date]:
    value = str(raw or "").strip()
    if not value:
        return None
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(value[:10], fmt).date()
        except Exception:
            continue
    return None


def _quarter_period_from_date(raw: str, today: Optional[date] = None) -> Optional[int]:
    dt = _parse_registry_date(raw)
    if dt is None:
        return None
    now = today or date.today()
    period = (now.year - dt.year) * 4 + ((now.month - 1) // 3) - ((dt.month - 1) // 3) + 1
    return max(1, int(period))


def apply_reduction_coefficients(minimum: float, user: UserFlags, use_rar: bool) -> float:
    if not reductions_allowed(user):
        return float(minimum)
    m = float(minimum)
    m *= get_general_reduction_coefficient(user, use_rar)
    m *= get_new_contract_coefficient(user)
    return m


def calculate_special_minimum_by_total_population(
    licenses: List[RadioLicense],
    contract_rate: float,
    use_rar: bool,
    is_vois: bool,
) -> float:
    total_pop = sum(int(lic.population or 0) for lic in licenses)
    row = _pick_min_row(contract_rate, use_rar, is_vois)
    base_minimum = float(row[get_population_bucket_index(total_pop)])
    valid_hours = [float(lic.weekly_hours() or 0.0) for lic in licenses if float(lic.weekly_hours() or 0.0) > 0]
    if valid_hours:
        avg_hours = sum(valid_hours) / len(valid_hours)
        base_minimum *= get_hours_coefficient(avg_hours)
    return float(base_minimum)


def calculate_contract_minimum(
    licenses: List[RadioLicense],
    user: UserFlags,
    contract_rate: float,
    use_rar: bool,
    is_vois: bool,
) -> Tuple[float, List[str], List[Dict[str, Any]], str]:
    parts: List[float] = []
    details: List[Dict[str, Any]] = []
    lines: List[str] = ["Минимальные суммы по лицензиям:"]
    for lic in licenses:
        m = calculate_license_minimum(lic, user, use_rar, is_vois)
        offline_m = calculate_license_minimum_without_internet(lic, user, use_rar, is_vois)
        internet_addon = max(0.0, float(m) - float(offline_m))
        own_site_count = max(
            [max(0, int(ch.site_app_count or 0)) for ch in (lic.channels or []) if ch.founder_type == "user"]
            or [0]
        )
        parts.append(float(m))
        details.append({
            "license": lic,
            "base_minimum": float(m),
            "internet_addon": float(internet_addon),
            "internet_income_threshold_label": (
                format_internet_income_threshold_label(user.previous_year_income_or_expense)
                if internet_addon > 0 and own_site_count > 0
                else ""
            ),
            "license_rate": float(calculate_license_rate(lic, use_rar, is_vois)),
        })
        lines.append(
            f"• {lic.license_id}: ставка {calculate_license_rate(lic, use_rar, is_vois):.1f}%, "
            f"население {format_count(lic.population)}, диапазон {_population_bucket_label(lic.population)} → {money(m)} ₽."
        )
    formula = " + ".join(money(x) for x in parts) or "0"
    total_by_licenses = float(sum(parts))
    lines.append(f"Сумма по лицензиям: {formula} = {money(total_by_licenses)} ₽.")

    has_third_party_scenario = bool(user.has_third_party_channels) or any(
        ch.founder_type == "third_party"
        for lic in licenses
        for ch in (lic.channels or [])
    )
    if has_third_party_scenario:
        lines.append(
            "Применён сценарий смешанного вещания (свои и/или сторонние радиоканалы на одной частоте): "
            "минимальная сумма определяется суммированием по лицензиям/каналам."
        )
        lines.append(f"Итого минимальная сумма: {money(total_by_licenses)} ₽.")
        return total_by_licenses, lines, details, "license_sum"

    lines.append(f"Итого минимальная сумма: {money(total_by_licenses)} ₽.")
    return total_by_licenses, lines, details, "license_sum"


def apply_10_percent_rule(
    licenses: List[RadioLicense],
    contract_rate: float,
    minimum: float,
    percent_fee: float,
    user: UserFlags,
    use_rar: bool,
    is_vois: bool,
) -> float:
    if not user.has_documented_income:
        return float(minimum)
    if float(minimum) <= float(percent_fee) * 1.10:
        return float(minimum)
    if int(user.report_period_number_from_special_minimum_start or 1) > 8:
        return float(minimum)
    special_min = calculate_special_minimum_by_total_population(licenses, contract_rate, use_rar, is_vois)
    return float(special_min)


def _status_label(user: UserFlags) -> str:
    if str(user.status_code or "").strip() == "new_user":
        return "новый пользователь"
    if str(user.status_code or "").strip() == "existing_contract_new_licenses":
        return "есть действующий договор и новые лицензии"
    return "есть действующий договор"


def calculate_dynamic_minimum_by_licenses(
    minimum_details: List[Dict[str, Any]],
    user: UserFlags,
    use_rar: bool,
) -> Tuple[float, List[str], List[str], float, float]:
    general_coeff = get_general_reduction_coefficient(user, use_rar) if reductions_allowed(user) else 1.0
    lines: List[str] = ["Применение коэффициентов к минимальной сумме по лицензиям:"]
    applied_labels: List[str] = []
    if general_coeff != 1.0:
        if abs(general_coeff - 0.85) < 1e-9:
            applied_labels.append("0,85 (участник отраслевой организации / соглашения, в том числе участник РАР)")
        elif abs(general_coeff - 0.9) < 1e-9:
            applied_labels.append("0,9 (пакетное заключение)")
        elif abs(general_coeff - 0.95) < 1e-9:
            applied_labels.append("0,95 (несколько категорий использования)")
        else:
            applied_labels.append(str(general_coeff).replace(".", ","))
    parts_y1: List[float] = []
    parts_y2: List[float] = []
    parts_std: List[float] = []
    for item in minimum_details:
        lic: RadioLicense = item["license"]
        base_minimum = float(item["base_minimum"])
        current_std = base_minimum
        expr = [money(base_minimum)]
        if general_coeff != 1.0:
            current_std *= general_coeff
            expr.append(str(general_coeff).replace(".", ","))
        apply_new_coeff = bool(user.status_code == "new_user" or lic.license_id in user.new_license_ids)
        if apply_new_coeff:
            y1 = base.round_rub(current_std * 0.75)
            y2 = base.round_rub(current_std * 0.88)
            y9 = base.round_rub(current_std)
            parts_y1.append(float(y1))
            parts_y2.append(float(y2))
            parts_std.append(float(y9))
            if "0,75 (1–4 отчётные периоды)" not in applied_labels:
                applied_labels.append("0,75 (1–4 отчётные периоды)")
            if "0,88 (5–8 отчётные периоды)" not in applied_labels:
                applied_labels.append("0,88 (5–8 отчётные периоды)")
            lines.append(
                f"• {lic.license_id}: {' × '.join(expr)} → "
                f"1–4 период: {money(y1)} ₽; "
                f"5–8 период: {money(y2)} ₽; "
                f"с 9-го периода: {money(y9)} ₽."
            )
        else:
            final_part = base.round_rub(current_std)
            parts_y1.append(float(final_part))
            parts_y2.append(float(final_part))
            parts_std.append(float(final_part))
            lines.append(f"• {lic.license_id}: {' × '.join(expr)} = {money(final_part)} ₽.")
    total_y1 = float(sum(parts_y1))
    total_y2 = float(sum(parts_y2))
    total_std = float(sum(parts_std))
    lines.append(f"Итого 1–4 отчётные периоды: {' + '.join(money(x) for x in parts_y1)} = {money(total_y1)} ₽.")
    lines.append(f"Итого 5–8 отчётные периоды: {' + '.join(money(x) for x in parts_y2)} = {money(total_y2)} ₽.")
    lines.append(f"Итого с 9-го отчётного периода: {' + '.join(money(x) for x in parts_std)} = {money(total_std)} ₽.")
    return total_std, lines, applied_labels, total_y1, total_y2


def calculate_dynamic_minimum_from_selected_minimum(
    selected_minimum: float,
    user: UserFlags,
    use_rar: bool,
) -> Tuple[float, List[str], List[str], float, float]:
    general_coeff = get_general_reduction_coefficient(user, use_rar) if reductions_allowed(user) else 1.0
    applied_labels: List[str] = []
    if general_coeff != 1.0:
        if abs(general_coeff - 0.85) < 1e-9:
            applied_labels.append("0,85 (участник отраслевой организации / соглашения, в том числе участник РАР)")
        elif abs(general_coeff - 0.9) < 1e-9:
            applied_labels.append("0,9 (пакетное заключение)")
        elif abs(general_coeff - 0.95) < 1e-9:
            applied_labels.append("0,95 (несколько категорий использования)")
        else:
            applied_labels.append(str(general_coeff).replace(".", ","))

    base_selected = float(selected_minimum)
    y9 = base.round_rub(base_selected * general_coeff)
    y1 = base.round_rub(float(y9) * 0.75)
    y2 = base.round_rub(float(y9) * 0.88)
    applied_labels.extend(["0,75 (1–4 отчётные периоды)", "0,88 (5–8 отчётные периоды)"])
    lines = [
        "Применение коэффициентов к выбранной минимальной сумме:",
        f"• Выбранная минимальная сумма: {money(selected_minimum)} ₽.",
        (
            f"• 1–4 отчётные периоды: {money(y1)} ₽; "
            f"5–8 отчётные периоды: {money(y2)} ₽; "
            f"с 9-го отчётного периода: {money(y9)} ₽."
        ),
    ]
    return float(y9), lines, applied_labels, float(y1), float(y2)


def _filter_and_dedupe_notes(notes: List[str], selected_ids: Set[str]) -> List[str]:
    out: List[str] = []
    seen: Set[str] = set()
    for note in notes:
        text = str(note or "").strip()
        if not text or text in seen:
            continue
        m = re.search(r"Лицензия\s+(Л[0-9А-ЯA-Z\-\/]+)", text)
        if m and selected_ids and m.group(1).strip() not in selected_ids:
            continue
        out.append(text)
        seen.add(text)
    return out


def _infer_founder_type(ch_name: str, has_third_party_channels: bool) -> str:
    if not has_third_party_channels:
        return "user"
    # Базовое правило: если включён признак смешанного вещания, считаем первый канал own, остальные third_party.
    # Детальная ручная раскладка может быть задана через --channel_founder.
    return "third_party"


def parse_channel_founders(items: Optional[List[str]]) -> Dict[str, str]:
    out: Dict[str, str] = {}
    if not items:
        return out
    for raw in items:
        if not raw:
            continue
        parts = str(raw).split("=", 1)
        if len(parts) != 2:
            continue
        key = parts[0].strip()
        val = parts[1].strip().lower()
        if not key or val not in {"user", "third_party"}:
            continue
        out[key] = val
    return out


def _parse_int_by_license(items: Optional[List[str]]) -> Dict[str, int]:
    out: Dict[str, int] = {}
    if not items:
        return out
    for raw in items:
        if not raw:
            continue
        parts = str(raw).split("=", 1)
        if len(parts) != 2:
            continue
        lic = parts[0].strip()
        if not lic:
            continue
        try:
            val = int(float(parts[1].strip().replace(",", ".")))
        except Exception:
            continue
        out[lic] = max(0, val)
    return out


def _parse_sites_by_license(items: Optional[List[str]]) -> Dict[str, List[str]]:
    out: Dict[str, List[str]] = {}
    if not items:
        return out
    for raw in items:
        if not raw:
            continue
        parts = str(raw).split("=", 1)
        if len(parts) != 2:
            continue
        lic = parts[0].strip()
        if not lic:
            continue
        sites = [x.strip() for x in str(parts[1] or "").split("|") if x.strip()]
        out[lic] = sites
    return out


def _prepare_channels_by_user_flags(
    licenses: List[RadioLicense],
    user: UserFlags,
    channel_founders: Dict[str, str],
    actual_share_by_channel: Dict[str, float],
    internet_resources_by_license: Dict[str, int],
) -> None:
    for lic in licenses:
        lic_site_count = int(internet_resources_by_license.get(lic.license_id, 0))
        if lic_site_count <= 0 and bool(user.simultaneous_internet_broadcast):
            lic_site_count = max(0, int(user.site_app_count or 0))
        for i, ch in enumerate(lic.channels):
            key1 = f"{lic.license_id}|{ch.name}"
            key2 = f"{lic.license_id}|{i+1}"
            ch.founder_type = channel_founders.get(key1) or channel_founders.get(key2) or ("user" if i == 0 else _infer_founder_type(ch.name, user.has_third_party_channels))
            ch.simultaneous_internet_broadcast = bool(lic_site_count > 0 or user.simultaneous_internet_broadcast)
            ch.internet_admin_by_user = bool(lic_site_count > 0 or user.internet_admin_by_user)
            # Количество ресурсов задаётся на лицензию, не на каждый канал.
            ch.site_app_count = int(lic_site_count if i == 0 else 0)
            if key1 in actual_share_by_channel:
                ch.actual_share_percent = float(actual_share_by_channel[key1])
            elif key2 in actual_share_by_channel:
                ch.actual_share_percent = float(actual_share_by_channel[key2])


def parse_actual_share_by_channel(items: Optional[List[str]]) -> Dict[str, float]:
    out: Dict[str, float] = {}
    if not items:
        return out
    for raw in items:
        if not raw:
            continue
        parts = str(raw).split("=", 1)
        if len(parts) != 2:
            continue
        key = parts[0].strip()
        try:
            val = float(parts[1].strip().replace(",", "."))
        except Exception:
            continue
        if not key:
            continue
        if val < 0 or val > 100:
            continue
        out[key] = val
    return out


def determine_income_base(args, user: UserFlags) -> Tuple[Optional[float], List[str]]:
    notes: List[str] = []
    base_type_label = format_financial_base_type(getattr(args, "base_type", None))
    if args.revenue_q is not None:
        notes.append(f"Финансовая база: {base_type_label if base_type_label != 'не указана' else 'доходы за квартал'}")
        return float(args.revenue_q), notes
    if args.annual_revenue is not None:
        notes.append(f"Финансовая база: {base_type_label if base_type_label != 'не указана' else 'годовая выручка или доход'}")
        return float(args.annual_revenue) / 4.0, notes
    if user.is_100_state_capital and args.expenses_q is not None:
        notes.append(f"Финансовая база: {base_type_label if base_type_label != 'не указана' else 'расходы за квартал'}")
        return float(args.expenses_q), notes
    if args.expenses_q is not None:
        notes.append(f"Финансовая база: {base_type_label if base_type_label != 'не указана' else 'расходы за квартал'}")
        return float(args.expenses_q), notes
    return None, notes


def money(x: Optional[float]) -> str:
    if x is None:
        return "—"
    return f"{float(x):,.2f}".replace(",", " ").replace(".00", "")


def money_precise(x: Optional[Any]) -> str:
    """Без округления: для расчётных сумм по процентной ставке."""
    if x is None:
        return "—"
    d = Decimal(str(x))
    s = format(d, "f")
    if "." in s:
        s = s.rstrip("0").rstrip(".")
    sign = "-" if s.startswith("-") else ""
    if sign:
        s = s[1:]
    int_part, dot, frac_part = s.partition(".")
    int_fmt = f"{int(int_part or '0'):,}".replace(",", " ")
    return f"{sign}{int_fmt}{('.' + frac_part) if dot and frac_part else ''}"


def format_count(x: Optional[Any]) -> str:
    if x is None:
        return "—"
    d = Decimal(str(x))
    sign = "-" if d < 0 else ""
    d = abs(d)
    if d == d.to_integral_value():
        return f"{sign}{int(d):,}".replace(",", " ")
    s = format(d.normalize(), "f")
    int_part, dot, frac_part = s.partition(".")
    int_fmt = f"{int(int_part or '0'):,}".replace(",", " ")
    frac_part = frac_part.rstrip("0")
    return f"{sign}{int_fmt}{('.' + frac_part) if dot and frac_part else ''}"


def format_percent(x: Optional[Any]) -> str:
    if x is None:
        return "—"
    s = f"{float(x):.1f}".replace(".", ",")
    if s.endswith(",0"):
        s = s[:-2]
    return f"{s}%"


def _h(value: Any) -> str:
    return html_lib.escape("" if value is None else str(value), quote=True)


def _rkn_license_url(license_id: str) -> str:
    from urllib.parse import quote

    return "https://rkn.gov.ru/activity/mass-media/for-broadcasters/teleradio/?id=" + quote(str(license_id), safe="")


def _license_link(license_id: str) -> str:
    value = str(license_id or "").strip()
    if not value:
        return "—"
    return (
        f'<a class="radioReportLicenseLink" href="{_h(_rkn_license_url(value))}" '
        f'target="_blank" rel="noopener noreferrer">{_h(value)}</a>'
    )


def _is_report_result_label(label: Any) -> bool:
    normalized = str(label or "").strip().lower()
    return normalized in {
        "минимальная сумма в квартал по договору",
        "минимальная сумма в квартал по суммарной численности населения",
        "процентная ставка по договору",
    }


def _report_rows(rows: List[Tuple[str, Any]]) -> str:
    body = []
    for label, value in rows:
        if value is None or value == "":
            continue
        label_html: Any = _h(label)
        marker_class = " radioReportResultRow" if _is_report_result_label(label) and str(value).strip() != "—" else ""
        body.append(
            f'<div class="radioReportRow{marker_class}">'
            f'<div class="radioReportKey">{label_html}</div>'
            f'<div class="radioReportValue">{value if isinstance(value, _SafeHtml) else _h(value)}</div>'
            '</div>'
        )
    return '<div class="radioReportRows">' + "".join(body) + "</div>"


def _report_facts(rows: List[Tuple[str, Any]]) -> str:
    body = []
    for label, value in rows:
        if value is None or value == "":
            continue
        label_html: Any = _h(label)
        marker_class = " radioReportResultFact" if _is_report_result_label(label) and str(value).strip() != "—" else ""
        rendered = value if isinstance(value, _SafeHtml) else _h(value)
        body.append(
            f'<p class="radioReportFact{marker_class}">'
            f'<strong>{label_html}:</strong> {rendered}'
            '</p>'
        )
    return '<div class="radioReportFacts">' + "".join(body) + "</div>"


def _terms_block(title: str, body: str) -> str:
    content = body.strip() if str(body or "").strip() else "<p>—</p>"
    return (
        '<div class="radioReportTermsBlock">'
        f'<h4>{_h(title)}</h4>'
        f'{content}'
        '</div>'
    )


def _format_percent_minimum_terms(rate: Optional[float], minimum: Optional[float]) -> str:
    if rate is None or minimum is None:
        return "<p>—</p>"
    return (
        f"<p>{format_percent(rate)} от доходов/расходов, "
        f"но не менее {money(minimum)} ₽ за квартал.</p>"
    )


class _SafeHtml(str):
    pass


def _safe_html(value: str) -> _SafeHtml:
    return _SafeHtml(value)


def _section(title: str, body: str) -> str:
    return (
        '<section class="radioReportSection">'
        f'<h3>{_h(title)}</h3>'
        f'{body}'
        '</section>'
    )


def _clean_note(note: str) -> str:
    text = str(note or "").strip()
    text = re.sub(r"\bmin\.s\.\s*=", "Расчёт минимальной суммы:", text, flags=re.I)
    text = text.replace(" х 088", " × 0,88").replace(" x 088", " × 0,88")
    text = text.replace("х 088", "× 0,88").replace("x 088", "× 0,88")
    text = text.replace("50 ₽ 000 000", "50 000 000 ₽")
    return text


def _license_rate_rows(lic: RadioLicense, use_rar: bool, is_vois: bool, is_new: bool) -> str:
    heading = "Новая лицензия" if is_new else "Лицензия"
    heading_html = '<span class="radioReportMarker">Новая лицензия</span>' if is_new else _h(heading)
    license_rate = calculate_license_rate(lic, use_rar, is_vois)
    parts = [
        f'<div class="radioReportLicenseTitle">{heading_html} № {_license_link(lic.license_id)} — {format_percent(license_rate)}</div>'
    ]
    channel_rows: List[str] = []
    rate_table = _rate_table(use_rar, is_vois)
    for ch in lic.channels or []:
        cr = calculate_channel_rate(ch, use_rar, is_vois)
        channel_rows.append(
            '<div class="radioReportChannel">'
            f'<div class="radioReportChannelTitle">{_h(ch.name)} — {format_count(ch.weekly_hours)} ч — {format_percent(cr)}</div>'
        )
        if ch.actual_share_percent is not None:
            channel_rows.append(
                f'<div class="radioReportMuted">Фактическая доля использования: '
                f'{format_percent(ch.actual_share_percent)}. Ставка: {format_percent(cr)}.</div>'
            )
        elif ch.topics:
            seen_topics: Set[str] = set()
            topic_items: List[str] = []
            for t in ch.topics:
                norm_topic = base.normalize_topic(t.name)
                if norm_topic in seen_topics:
                    continue
                seen_topics.add(norm_topic)
                topic_rate = rate_table.get(t.category, rate_table["UNKNOWN"])
                topic_items.append(
                    f'<li>{_h(t.name)} → категория {_h(t.category)} — {format_percent(topic_rate)}</li>'
                )
            if topic_items:
                channel_rows.append('<ul class="radioReportList">' + "".join(topic_items) + "</ul>")
        channel_rows.append("</div>")
    if not channel_rows:
        channel_rows.append('<div class="radioReportMuted">СМИ/радиоканалы не указаны.</div>')
    parts.append("".join(channel_rows))
    return '<div class="radioReportLicenseBlock">' + "".join(parts) + "</div>"


def _license_minimum_rows(
    lic: RadioLicense,
    user: UserFlags,
    use_rar: bool,
    is_vois: bool,
    internet_resources_by_license: Dict[str, int],
    minimum_details_by_id: Dict[str, Dict[str, Any]],
) -> str:
    details = minimum_details_by_id.get(lic.license_id) or {}
    minimum_value = details.get("base_minimum")
    internet_addon = float(details.get("internet_addon") or 0.0)
    hours = lic.weekly_hours()
    hours_coeff = get_hours_coefficient(hours)
    internet_count = int(internet_resources_by_license.get(lic.license_id, 0))
    rows: List[Tuple[str, Any]] = [
        ("Численность населения", format_count(lic.population)),
        ("Диапазон по численности населения", _population_bucket_label(lic.population)),
        ("Интернет-ресурсы", format_count(internet_count)),
        ("Всего часов вещания", format_count(hours)),
        ("Коэффициент за объём вещания", "не применяется" if abs(hours_coeff - 1.0) < 1e-9 else str(hours_coeff).replace(".", ",")),
    ]
    internet_label = format_internet_component_label(internet_count)
    internet_line = (
        f'<div class="radioReportLicenseAddon">{internet_label}: {money(internet_addon)} ₽.</div>'
        if internet_addon > 0
        else ""
    )
    heading = "Новая лицензия" if lic.license_id in user.new_license_ids else "Лицензия"
    heading_html = '<span class="radioReportMarker">Новая лицензия</span>' if heading == "Новая лицензия" else _h(heading)
    return (
        '<div class="radioReportLicenseBlock">'
        f'<div class="radioReportLicenseTitle radioReportMinimumTitle">{heading_html} № {_license_link(lic.license_id)} — минимальная сумма: {money(minimum_value)} ₽</div>'
        f'{internet_line}'
        f'{_report_facts(rows)}'
        '</div>'
    )


def _build_report(
    inn: str,
    licenses: List[RadioLicense],
    use_rar: bool,
    is_vois: bool,
    standard_contract_rate: Optional[float],
    contract_rate: Optional[float],
    income_base_q: Optional[float],
    base_percent_fee: Optional[float],
    standard_minimum_fee: Optional[float],
    base_minimum_fee: Optional[float],
    final_percent_fee: Optional[float],
    final_minimum_fee: Optional[float],
    total_population_minimum_fee: Optional[float],
    addendum_percent_fee: Optional[float],
    addendum_minimum_fee: Optional[float],
    min_with_coeff_y1: Optional[float],
    min_with_coeff_y2: Optional[float],
    addendum_min_with_coeff_y1: Optional[float],
    addendum_min_with_coeff_y2: Optional[float],
    payable: Optional[float],
    calc_mode: str,
    user_discount_label: str,
    new_user: bool,
    contract_period: int,
    guillotine_triggered: bool,
    applied_coeffs: List[str],
    volume_coeff_notes: List[str],
    internet_resources_by_license: Dict[str, int],
    internet_sites_by_license: Dict[str, List[str]],
    minimum_breakdown: List[str],
    minimum_details: List[Dict[str, Any]],
    minimum_source: str,
    notes: List[str],
    needs: List[str],
    status_label: str,
    show_global_new_user_terms: bool,
    user: UserFlags,
) -> str:
    org = licenses[0].org_name if licenses else "Организация не найдена"
    total_internet = sum(max(0, int(v or 0)) for v in internet_resources_by_license.values())
    minimum_details_by_id = {
        str(item.get("license").license_id): item
        for item in minimum_details
        if item.get("license") is not None
    }
    has_volume_coeff = bool(volume_coeff_notes)
    minimum_exceeds_percent = (
        base_minimum_fee is not None
        and base_percent_fee is not None
        and float(base_minimum_fee) > float(base_percent_fee)
    )

    source_rows = [
        ReportRow("Наименование", org),
        ReportRow("ИНН", inn),
        ReportRow("Финансовая база за квартал", f"{money(income_base_q)} ₽" if income_base_q is not None else "не указана"),
        ReportRow("Статус пользователя", status_label),
        ReportRow("Фиксированная сумма вознаграждения", "да" if calc_mode == "fixed" else "нет"),
        ReportRow("Понижающий коэффициент", user_discount_label or "не применяется"),
        ReportRow("Количество лицензий, участвующих в расчёте", format_count(len(licenses))),
        ReportRow("Количество интернет-ресурсов", format_count(total_internet)),
    ]

    rate_table = _rate_table(use_rar, is_vois)
    rate_licenses: List[RateLicense] = []
    for lic in licenses:
        rate_channels: List[RateChannel] = []
        for ch in lic.channels or []:
            channel_rate = calculate_channel_rate(ch, use_rar, is_vois)
            topic_rows: List[RateTopic] = []
            if ch.actual_share_percent is None:
                seen_topics: Set[str] = set()
                for topic in ch.topics:
                    norm_topic = base.normalize_topic(topic.name)
                    if norm_topic in seen_topics:
                        continue
                    seen_topics.add(norm_topic)
                    topic_rows.append(
                        RateTopic(
                            name=topic.name,
                            category=topic.category,
                            rate=rate_table.get(topic.category, rate_table["UNKNOWN"]),
                        )
                    )
            rate_channels.append(
                RateChannel(
                    name=ch.name,
                    weekly_hours=float(ch.weekly_hours or 0.0),
                    rate=channel_rate,
                    actual_share_percent=ch.actual_share_percent,
                    topics=topic_rows,
                )
            )
        rate_licenses.append(
            RateLicense(
                license_number=lic.license_id,
                license_url=lic.rkn_url or _rkn_license_url(lic.license_id),
                is_new=lic.license_id in user.new_license_ids,
                rate=calculate_license_rate(lic, use_rar, is_vois),
                channels=rate_channels,
            )
        )

    minimum_licenses: List[MinimumLicense] = []
    for lic in licenses:
        details = minimum_details_by_id.get(lic.license_id) or {}
        minimum_licenses.append(
            MinimumLicense(
                license_number=lic.license_id,
                license_url=lic.rkn_url or _rkn_license_url(lic.license_id),
                is_new=lic.license_id in user.new_license_ids,
                minimum=details.get("base_minimum"),
                internet_component=float(details.get("internet_addon") or 0.0),
                internet_resources=int(internet_resources_by_license.get(lic.license_id, 0)),
                population=int(lic.population) if lic.population is not None else None,
                population_range=_population_bucket_label(lic.population),
                weekly_hours=float(lic.weekly_hours() or 0.0),
                hours_coefficient=get_hours_coefficient(lic.weekly_hours()),
            )
        )

    total_pop = sum(int(lic.population or 0) for lic in licenses)
    show_total_population_minimum = total_population_minimum_fee is not None
    minimum_rows: List[ReportRow] = [
        ReportRow(
            "Минимальная сумма в квартал по договору",
            f"{money(final_minimum_fee)} ₽" if final_minimum_fee is not None else "—",
            True,
        ),
        ReportRow("Минимальная сумма превышает расчётную сумму", "да" if minimum_exceeds_percent else "нет"),
        ReportRow(
            "Минимальная сумма в квартал по суммарной численности населения",
            f"{money(total_population_minimum_fee)} ₽" if total_population_minimum_fee is not None else "—",
            total_population_minimum_fee is not None,
        ),
    ]
    if show_total_population_minimum and total_population_minimum_fee is not None:
        minimum_rows.append(ReportRow("Суммарная численность населения", format_count(total_pop)))
    if has_volume_coeff or applied_coeffs:
        minimum_rows.append(ReportRow("Примечание", "Минимальная сумма указана с учётом применимых коэффициентов."))

    def percent_minimum_terms(rate: Optional[float], minimum: Optional[float]) -> List[str]:
        if rate is None or minimum is None:
            return ["—"]
        return [f"{format_percent(rate)} от доходов/расходов, но не менее {money(minimum)} ₽ за квартал."]

    standard_terms = ContractTerms(
        title="Стандартные условия",
        lines=percent_minimum_terms(standard_contract_rate, standard_minimum_fee),
    )
    financial_terms = ContractTerms(title="Финансовые условия с учётом коэффициентов")
    effective_terms_differ = bool(
        (contract_rate is not None and standard_contract_rate is not None and abs(float(contract_rate) - float(standard_contract_rate)) > 1e-9)
        or (
            final_minimum_fee is not None
            and standard_minimum_fee is not None
            and abs(float(final_minimum_fee) - float(standard_minimum_fee)) > 1e-9
        )
    )
    has_financial_adjustments = bool(
        applied_coeffs
        or show_global_new_user_terms
        or volume_coeff_notes
        or effective_terms_differ
    )
    if has_financial_adjustments:
        if show_global_new_user_terms and (min_with_coeff_y1 is not None or min_with_coeff_y2 is not None):
            if min_with_coeff_y1 is not None:
                financial_terms.bullets.append(
                    f"С 1-го по 4-й отчётный период: {format_percent(contract_rate)} от доходов/расходов, "
                    f"но не менее {money(min_with_coeff_y1)} ₽ за квартал."
                )
            if min_with_coeff_y2 is not None:
                financial_terms.bullets.append(
                    f"С 5-го по 8-й отчётный период: {format_percent(contract_rate)} от доходов/расходов, "
                    f"но не менее {money(min_with_coeff_y2)} ₽ за квартал."
                )
            if final_minimum_fee is not None:
                financial_terms.bullets.append(
                    f"С 9-го отчётного периода: {format_percent(contract_rate)} от доходов/расходов, "
                    f"но не менее {money(final_minimum_fee)} ₽ за квартал."
                )
        else:
            financial_terms.lines.extend(percent_minimum_terms(contract_rate, final_minimum_fee))
        if volume_coeff_notes:
            for note in volume_coeff_notes:
                financial_terms.bullets.append(_clean_note(note))
    else:
        financial_terms.lines.append("—")

    addendum_applicable = bool((guillotine_triggered and addendum_minimum_fee is not None) or calc_mode == "fixed")
    addendum_terms = ContractTerms(title="Условия в дополнительном соглашении")
    if addendum_applicable:
        if calc_mode == "fixed" and payable is not None:
            addendum_terms.lines.append(f"Фиксированная сумма вознаграждения: {money(payable)} ₽ за квартал.")
        if addendum_percent_fee is not None:
            addendum_terms.lines.append(f"Расчётная сумма по договору за квартал: {money_precise(addendum_percent_fee)} ₽.")
        if contract_rate is not None and addendum_minimum_fee is not None:
            if show_global_new_user_terms and (addendum_min_with_coeff_y1 is not None or addendum_min_with_coeff_y2 is not None):
                if addendum_min_with_coeff_y1 is not None:
                    addendum_terms.bullets.append(
                        f"С 1-го по 4-й отчётный период: {format_percent(contract_rate)} от доходов/расходов, "
                        f"но не менее {money(addendum_min_with_coeff_y1)} ₽ за квартал."
                    )
                if addendum_min_with_coeff_y2 is not None:
                    addendum_terms.bullets.append(
                        f"С 5-го по 8-й отчётный период: {format_percent(contract_rate)} от доходов/расходов, "
                        f"но не менее {money(addendum_min_with_coeff_y2)} ₽ за квартал."
                    )
                addendum_terms.bullets.append(
                    f"С 9-го отчётного периода: {format_percent(contract_rate)} от доходов/расходов, "
                    f"но не менее {money(addendum_minimum_fee)} ₽ за квартал."
                )
            else:
                addendum_terms.lines.append(
                    f"{format_percent(contract_rate)} от доходов/расходов, "
                    f"но не менее {money(addendum_minimum_fee)} ₽ за квартал."
                )
        if guillotine_triggered:
            addendum_terms.lines.append(
                "Условие по минимальной сумме по суммарной численности населения действует не более восьми полных отчётных периодов."
            )
    if not addendum_terms.lines and not addendum_terms.bullets:
        addendum_terms.lines.append("—")

    comment_items: List[str] = []
    for note in notes:
        cleaned = _clean_note(note)
        if cleaned:
            comment_items.append(cleaned)
    for need in needs:
        cleaned = _clean_note(need)
        if cleaned:
            comment_items.append(f"Нужно уточнить: {cleaned}")

    model = RadioReportModel(
        title="Результат расчёта",
        source_data=source_rows,
        rate_licenses=rate_licenses,
        contract_rate=contract_rate,
        quarter_amount=float(base_percent_fee) if base_percent_fee is not None else None,
        minimum_licenses=minimum_licenses,
        minimum_rows=minimum_rows,
        contract_terms=[standard_terms, financial_terms, addendum_terms],
        comments=comment_items,
    )
    return render_radio_report_html(model, embed_model=True)


# --------------------------- main ---------------------------

def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--inn", required=True)
    ap.add_argument("--rkn_xlsx", required=True)
    ap.add_argument("--vars_xlsx", required=True)

    ap.add_argument("--annual_revenue", type=float, default=None)
    ap.add_argument("--revenue_q", type=float, default=None)
    ap.add_argument("--expenses_q", type=float, default=None)
    ap.add_argument("--base_type", type=str, default=None)
    ap.add_argument("--licenses", action="append", default=None)

    ap.add_argument("--calc_mode", type=str, default="percent_minimum", choices=["percent_minimum", "fixed"])
    ap.add_argument("--is_rar_member", action="store_true")
    ap.add_argument("--is_state_budget_institution", action="store_true")
    ap.add_argument("--is_internet_only", action="store_true")
    ap.add_argument("--has_documented_income", action="store_true")
    ap.add_argument("--is_100_state_capital", action="store_true")
    ap.add_argument("--new_user", action="store_true")
    ap.add_argument("--contract_period_number", type=int, default=1)
    ap.add_argument("--special_min_period_number", type=int, default=1)
    ap.add_argument("--is_new_user_and_other_use_contract", action="store_true")
    ap.add_argument("--is_package_contract", action="store_true")
    ap.add_argument("--assoc_member", action="store_true")
    ap.add_argument("--has_legal_cases", action="store_true")
    ap.add_argument("--signed_within_30_days", action="store_true")
    ap.add_argument("--has_current_agreement_breach", action="store_true")
    ap.add_argument("--previous_year_income_or_expense", type=float, default=0.0)
    ap.add_argument("--fixed_fee_eligible", action="store_true")
    ap.add_argument("--fixed_fee_basis_confirmed", action="store_true")
    ap.add_argument("--market_index", type=float, default=1.0)
    ap.add_argument("--inflation_index", type=float, default=1.0)

    ap.add_argument("--simultaneous_internet_broadcast", action="store_true")
    ap.add_argument("--internet_admin_by_user", action="store_true")
    ap.add_argument("--site_app_count", type=int, default=0)
    ap.add_argument("--internet_resources_by_license", action="append", default=None, help="LICENSE=COUNT")
    ap.add_argument("--internet_sites_by_license", action="append", default=None, help="LICENSE=site1|site2")
    ap.add_argument("--has_third_party_channels", action="store_true")
    ap.add_argument("--channel_founder", action="append", default=None, help="LICENSE|CHANNEL=user|third_party")
    ap.add_argument("--actual_share_by_channel", action="append", default=None, help="LICENSE|CHANNEL=87.5")
    ap.add_argument("--new_license_ids", action="append", default=None, help="LICENSE_ID для новых лицензий")
    ap.add_argument("--non_interactive", action="store_true")
    ap.add_argument("--skip_runtime_population_normalization", action="store_true")
    ap.add_argument("--society", type=str, default="РАО")
    args = ap.parse_args(argv)

    try:
        inn = base.parse_inn(args.inn)
    except Exception as e:
        print(f"Ошибка: {e}")
        return 2

    rkn_xlsx = Path(args.rkn_xlsx)
    vars_xlsx = Path(args.vars_xlsx)
    notes: List[str] = []
    needs: List[str] = []

    licenses, load_notes = load_radio_licenses_by_inn(
        rkn_xlsx=rkn_xlsx,
        inn=inn,
        vars_xlsx=vars_xlsx,
        runtime_population_normalization=not bool(args.skip_runtime_population_normalization),
    )
    notes.extend(load_notes)

    selected_ids = base.parse_license_list(args.licenses, None)
    selected_set: Set[str] = set()
    if selected_ids:
        selected_set = {str(x).strip() for x in selected_ids}
        licenses = [x for x in licenses if str(x.license_id).strip() in selected_set]
    if not licenses:
        print("Ошибка: по ИНН не найдено лицензий на осуществление радиовещания для расчёта.")
        return 2
    if not selected_set:
        selected_set = {str(x.license_id).strip() for x in licenses}
    notes = _filter_and_dedupe_notes(notes, selected_set)
    new_license_ids = {str(x).strip() for x in (base.parse_license_list(args.new_license_ids, None) or []) if str(x).strip()}
    if new_license_ids:
        selected_set = {str(x.license_id).strip() for x in licenses}
        new_license_ids = {x for x in new_license_ids if x in selected_set}
    status_code = "new_user" if bool(args.new_user) else ("existing_contract_new_licenses" if new_license_ids else "existing_contract")

    user = UserFlags(
        status_code=status_code,
        new_license_ids=new_license_ids,
        is_rar_member=bool(args.is_rar_member),
        is_state_budget_institution=bool(args.is_state_budget_institution),
        is_internet_only=bool(args.is_internet_only),
        has_documented_income=bool(args.has_documented_income),
        is_100_state_capital=bool(args.is_100_state_capital),
        is_new_user_initiated_contract=bool(args.new_user),
        report_period_number_from_contract_start=max(1, int(args.contract_period_number or 1)),
        is_new_user_and_other_use_contract=bool(args.is_new_user_and_other_use_contract),
        is_package_contract=bool(args.is_package_contract),
        assoc_member=bool(args.assoc_member),
        has_legal_cases=bool(args.has_legal_cases),
        signed_within_30_days=bool(args.signed_within_30_days),
        has_current_agreement_breach=bool(args.has_current_agreement_breach),
        report_period_number_from_special_minimum_start=max(1, int(args.special_min_period_number or 1)),
        previous_year_income_or_expense=float(args.previous_year_income_or_expense or 0.0),
        can_use_fixed_fee=bool(args.fixed_fee_eligible),
        fixed_fee_basis_confirmed=bool(args.fixed_fee_basis_confirmed),
        market_index=float(args.market_index or 1.0),
        inflation_index=float(args.inflation_index or 1.0),
        simultaneous_internet_broadcast=bool(args.simultaneous_internet_broadcast),
        internet_admin_by_user=bool(args.internet_admin_by_user),
        site_app_count=max(0, int(args.site_app_count or 0)),
        has_third_party_channels=bool(args.has_third_party_channels),
    )

    is_vois = _is_vois_society(args.society)
    use_rar = False if is_vois else use_rar_rates(user)
    if user.is_internet_only:
        text = _build_report(
            inn=inn,
            licenses=[],
            use_rar=use_rar,
            is_vois=is_vois,
            standard_contract_rate=None,
            contract_rate=None,
            income_base_q=None,
            base_percent_fee=None,
            standard_minimum_fee=None,
            base_minimum_fee=None,
            final_percent_fee=None,
            final_minimum_fee=None,
            total_population_minimum_fee=None,
            addendum_percent_fee=None,
            addendum_minimum_fee=None,
            min_with_coeff_y1=None,
            min_with_coeff_y2=None,
            addendum_min_with_coeff_y1=None,
            addendum_min_with_coeff_y2=None,
            payable=0.0,
            calc_mode=args.calc_mode,
            user_discount_label="не применяется",
            new_user=bool(args.new_user),
            contract_period=max(1, int(args.contract_period_number or 1)),
            guillotine_triggered=False,
            applied_coeffs=[],
            volume_coeff_notes=[],
            internet_resources_by_license={},
            internet_sites_by_license={},
            minimum_breakdown=[],
            minimum_details=[],
            minimum_source="license_sum",
            notes=["Положение для радиовещателей не применяется: internet-only вещание."],
            needs=[],
            status_label=_status_label(user),
            show_global_new_user_terms=False,
            user=user,
        )
        print(text)
        return 0

    channel_founders = parse_channel_founders(args.channel_founder)
    actual_share_by_channel = parse_actual_share_by_channel(args.actual_share_by_channel)
    internet_resources_by_license = _parse_int_by_license(args.internet_resources_by_license)
    internet_sites_by_license = _parse_sites_by_license(args.internet_sites_by_license)
    _prepare_channels_by_user_flags(
        licenses,
        user,
        channel_founders,
        actual_share_by_channel,
        internet_resources_by_license=internet_resources_by_license,
    )
    if any(ch.founder_type == "third_party" for lic in licenses for ch in (lic.channels or [])):
        user.has_third_party_channels = True

    contract_rate = calculate_contract_rate(licenses, use_rar, is_vois)
    standard_contract_rate = calculate_contract_rate(licenses, False, is_vois)
    income_base_q, base_notes = determine_income_base(args, user)
    notes.extend(base_notes)
    if income_base_q is None and args.calc_mode != "fixed":
        needs.append("Нужна база доходов/расходов за квартал или годовая база.")

    percent_fee = None
    if income_base_q is not None:
        # По запросу: расчётные суммы по процентной ставке НЕ округляем.
        percent_fee = (Decimal(str(income_base_q)) * Decimal(str(contract_rate))) / Decimal("100")
    base_percent_fee = percent_fee

    minimum_raw, minimum_breakdown, minimum_details, minimum_source = calculate_contract_minimum(
        licenses=licenses,
        user=user,
        contract_rate=contract_rate,
        use_rar=use_rar,
        is_vois=is_vois,
    )
    base_minimum_fee = base.round_rub(minimum_raw)
    standard_minimum_raw, _, _, _ = calculate_contract_minimum(
        licenses=licenses,
        user=user,
        contract_rate=standard_contract_rate,
        use_rar=False,
        is_vois=is_vois,
    )
    standard_minimum_fee = base.round_rub(standard_minimum_raw)
    for item in minimum_details:
        lic = item.get("license")
        internet_addon = float(item.get("internet_addon") or 0.0)
        threshold_label = str(item.get("internet_income_threshold_label") or "").strip()
        if lic is not None and internet_addon > 0 and threshold_label:
            notes.append(
                f"Лицензия {lic.license_id}, учтена минимальная сумма за интернет-вещание "
                f"{money(internet_addon)} ₽ ({threshold_label})"
            )
    volume_coeff_notes: List[str] = []
    for lic in licenses:
        h = float(lic.weekly_hours() or 0.0)
        hc = get_hours_coefficient(h)
        if hc < 1.0:
            volume_coeff_notes.append(
                f"{lic.license_id}: {h:g} ч/нед → коэффициент {str(hc).replace('.', ',')}"
            )

    k42 = get_general_reduction_coefficient(user, use_rar) if reductions_allowed(user) else 1.0
    user_discount_label = "не применяется"
    if k42 != 1.0:
        if abs(k42 - 0.85) < 1e-9:
            user_discount_label = "0,85 (участник отраслевой организации / соглашения, в том числе участник РАР)"
        elif abs(k42 - 0.9) < 1e-9:
            user_discount_label = "0,9 (пакетное заключение)"
        elif abs(k42 - 0.95) < 1e-9:
            user_discount_label = "0,95 (несколько категорий использования)"
        else:
            user_discount_label = str(k42).replace(".", ",")
    elif use_rar and user.assoc_member:
        user_discount_label = "0,85 (участник отраслевой организации / соглашения, в том числе участник РАР)"

    min_after_alt = base.round_rub(float(base_minimum_fee or 0.0) * k42)
    min_with_coeff_y1 = None
    min_with_coeff_y2 = None
    final_minimum_fee = min_after_alt
    dynamic_new_license_mode = bool(args.new_user or new_license_ids)
    show_global_new_user_terms = False
    dynamic_labels_for_summary: List[str] = []
    if dynamic_new_license_mode:
        if minimum_source == "license_sum":
            final_dynamic_minimum, dynamic_lines, dynamic_labels, dynamic_y1, dynamic_y2 = calculate_dynamic_minimum_by_licenses(
                minimum_details=minimum_details,
                user=user,
                use_rar=use_rar,
            )
        else:
            final_dynamic_minimum, dynamic_lines, dynamic_labels, dynamic_y1, dynamic_y2 = calculate_dynamic_minimum_from_selected_minimum(
                selected_minimum=float(base_minimum_fee or 0.0),
                user=user,
                use_rar=use_rar,
            )
        final_minimum_fee = base.round_rub(final_dynamic_minimum)
        min_with_coeff_y1 = base.round_rub(dynamic_y1)
        min_with_coeff_y2 = base.round_rub(dynamic_y2)
        show_global_new_user_terms = True
        minimum_breakdown.extend([""] + dynamic_lines)
        dynamic_labels_for_summary = list(dynamic_labels)
        for label in dynamic_labels:
            if label not in notes:
                notes.append(f"Применён коэффициент: {label}.")

    guillotine_triggered = False
    addendum_percent_fee = None
    addendum_minimum_fee = None
    addendum_min_with_coeff_y1 = None
    addendum_min_with_coeff_y2 = None
    total_population_minimum_fee = None
    if percent_fee is not None:
        base_minimum_for_3_3 = float(base_minimum_fee or 0.0)
        percent_fee_for_3_3 = float(percent_fee)
        base_minimum_decimal = Decimal(str(base_minimum_fee or 0))
        percent_fee_decimal = Decimal(str(percent_fee))
        ten_percent_threshold = (percent_fee_decimal * Decimal("1.10")).quantize(Decimal("0.01"))
        p33_triggered = bool(
            user.has_documented_income
            and base_minimum_decimal.quantize(Decimal("0.01")) > ten_percent_threshold
            and int(user.report_period_number_from_special_minimum_start or 1) <= 8
        )
        add_min = apply_10_percent_rule(
            licenses=licenses,
            contract_rate=contract_rate,
            minimum=base_minimum_for_3_3,
            percent_fee=percent_fee_for_3_3,
            user=user,
            use_rar=use_rar,
            is_vois=is_vois,
        )
        if p33_triggered:
            guillotine_triggered = True
            addendum_percent_fee = percent_fee
            add_base_after_alt = base.round_rub(float(add_min))
            addendum_minimum_fee = add_base_after_alt
            total_population_minimum_fee = add_base_after_alt
            if show_global_new_user_terms:
                addendum_min_with_coeff_y1 = base.round_rub(float(add_base_after_alt) * 0.75)
                addendum_min_with_coeff_y2 = base.round_rub(float(add_base_after_alt) * 0.88)
            notes.append("Пункт 3.3: применима минимальная сумма по суммарной численности населения.")

    payable: Optional[float]
    if args.calc_mode == "fixed":
        if not user.can_use_fixed_fee:
            notes.append("Режим фиксированной суммы запрошен, но признак допустимости не подтверждён; применён обычный расчёт.")
            args.calc_mode = "percent_minimum"
        elif not user.fixed_fee_basis_confirmed:
            args.calc_mode = "percent_minimum"
        else:
            annual_base = float(user.previous_year_income_or_expense or 0.0)
            if annual_base <= 0:
                needs.append("Для фиксированной суммы нужна база за предыдущий календарный год.")
                payable = None
            else:
                fixed = annual_base * float(contract_rate) / 100.0
                fixed *= float(user.market_index or 1.0)
                fixed *= float(user.inflation_index or 1.0)
                # Фикс не ниже минимальной суммы.
                fixed_q = max(fixed / 4.0, float(final_minimum_fee or 0.0))
                payable = base.round_rub(fixed_q)
                # Для расчётной суммы по ставке сохраняем точное значение без округления.
                base_percent_fee = Decimal(str(fixed_q))
                percent_fee = Decimal(str(fixed_q))
                final_minimum_fee = base.round_rub(fixed_q)
    if args.calc_mode != "fixed":
        if percent_fee is None:
            payable = final_minimum_fee if final_minimum_fee is not None else None
        else:
            payable = base.round_rub(max(float(percent_fee), float(final_minimum_fee or 0.0)))

    applied_coeffs: List[str] = []
    if k42 != 1.0:
        applied_coeffs.append(user_discount_label)
    for label in dynamic_labels_for_summary:
        if label not in applied_coeffs:
            applied_coeffs.append(label)

    text = _build_report(
        inn=inn,
        licenses=licenses,
        use_rar=use_rar,
        is_vois=is_vois,
        standard_contract_rate=standard_contract_rate,
        contract_rate=contract_rate,
        income_base_q=income_base_q,
        base_percent_fee=base_percent_fee,
        standard_minimum_fee=standard_minimum_fee,
        base_minimum_fee=base_minimum_fee,
        final_percent_fee=percent_fee,
        final_minimum_fee=final_minimum_fee,
        total_population_minimum_fee=total_population_minimum_fee,
        addendum_percent_fee=addendum_percent_fee,
        addendum_minimum_fee=addendum_minimum_fee,
        min_with_coeff_y1=min_with_coeff_y1,
        min_with_coeff_y2=min_with_coeff_y2,
        addendum_min_with_coeff_y1=addendum_min_with_coeff_y1,
        addendum_min_with_coeff_y2=addendum_min_with_coeff_y2,
        payable=payable,
        calc_mode=args.calc_mode,
        user_discount_label=user_discount_label,
        new_user=bool(args.new_user),
        contract_period=max(1, int(args.contract_period_number or 1)),
        guillotine_triggered=guillotine_triggered,
        applied_coeffs=applied_coeffs,
        volume_coeff_notes=volume_coeff_notes,
        internet_resources_by_license=internet_resources_by_license,
        internet_sites_by_license=internet_sites_by_license,
        minimum_breakdown=minimum_breakdown,
        minimum_details=minimum_details,
        minimum_source=minimum_source,
        notes=notes,
        needs=needs,
        status_label=_status_label(user),
        show_global_new_user_terms=show_global_new_user_terms,
        user=user,
    )
    print(text)
    return 0


def run_calc_capture(argv: List[str]) -> Tuple[int, str]:
    buf = io.StringIO()
    try:
        with redirect_stdout(buf):
            code = int(main(argv))
    except SystemExit as e:
        code = int(getattr(e, "code", 1) or 0)
    except Exception as e:
        code = 2
        buf.write(f"Ошибка: {type(e).__name__}: {e}\n")
    return code, buf.getvalue()
