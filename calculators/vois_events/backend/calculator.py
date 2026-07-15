"""VOIS events calculator entrypoint."""

from __future__ import annotations

from calculators.rao_events.backend.calculator import (
    EVENT_MUSIC_SHARE_COEFFICIENTS,
    EVENT_REGION_RATES,
    EventProfile,
    calculate_event_fee as _calculate_common_event_fee,
)


CALCULATOR_ID = "vois_events"
PROVIDER = "ВОИС"

VOIS_PROFILE = EventProfile(
    key="vois_events",
    provider="ВОИС",
    profile_title="Положение ВОИС №5 для трансляторов мероприятий",
    music_label="фонограмм",
    visits_point_label="2.4",
    free_access_point_label="2.3",
    rate_label_10k="Применяемая ставка вознаграждения (до 10 000 посещений)",
    rate_label_region="Применяемая ставка вознаграждения (согласно региону)",
    annual_fee_label="Сумма вознаграждения за год",
)


def calculate_event_fee(payload: dict):
    return _calculate_common_event_fee(payload, profile=VOIS_PROFILE)


__all__ = [
    "CALCULATOR_ID",
    "PROVIDER",
    "VOIS_PROFILE",
    "EVENT_REGION_RATES",
    "EVENT_MUSIC_SHARE_COEFFICIENTS",
    "calculate_event_fee",
]
