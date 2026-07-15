from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any, Optional


@dataclass
class ReportRow:
    label: str
    value: Any
    key_result: bool = False


@dataclass
class RateTopic:
    name: str
    category: str
    rate: Optional[float] = None


@dataclass
class RateChannel:
    name: str
    weekly_hours: Optional[float]
    rate: Optional[float]
    actual_share_percent: Optional[float] = None
    topics: list[RateTopic] = field(default_factory=list)


@dataclass
class RateLicense:
    license_number: str
    license_url: str
    is_new: bool
    rate: Optional[float]
    channels: list[RateChannel] = field(default_factory=list)


@dataclass
class MinimumLicense:
    license_number: str
    license_url: str
    is_new: bool
    minimum: Optional[float]
    internet_component: float = 0.0
    internet_resources: int = 0
    population: Optional[int] = None
    population_range: str = ""
    weekly_hours: Optional[float] = None
    hours_coefficient: Optional[float] = None


@dataclass
class ContractTerms:
    title: str
    lines: list[str] = field(default_factory=list)
    bullets: list[str] = field(default_factory=list)


@dataclass
class RadioReportModel:
    title: str
    source_data: list[ReportRow]
    rate_licenses: list[RateLicense]
    contract_rate: Optional[float]
    quarter_amount: Optional[float]
    minimum_licenses: list[MinimumLicense]
    minimum_rows: list[ReportRow]
    contract_terms: list[ContractTerms]
    comments: list[str]

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "RadioReportModel":
        def row(obj: dict[str, Any]) -> ReportRow:
            return ReportRow(**obj)

        def topic(obj: dict[str, Any]) -> RateTopic:
            return RateTopic(**obj)

        def channel(obj: dict[str, Any]) -> RateChannel:
            return RateChannel(
                name=obj.get("name", ""),
                weekly_hours=obj.get("weekly_hours"),
                rate=obj.get("rate"),
                actual_share_percent=obj.get("actual_share_percent"),
                topics=[topic(x) for x in obj.get("topics") or []],
            )

        def rate_license(obj: dict[str, Any]) -> RateLicense:
            return RateLicense(
                license_number=obj.get("license_number", ""),
                license_url=obj.get("license_url", ""),
                is_new=bool(obj.get("is_new")),
                rate=obj.get("rate"),
                channels=[channel(x) for x in obj.get("channels") or []],
            )

        def minimum_license(obj: dict[str, Any]) -> MinimumLicense:
            return MinimumLicense(**obj)

        def terms(obj: dict[str, Any]) -> ContractTerms:
            return ContractTerms(
                title=obj.get("title", ""),
                lines=list(obj.get("lines") or []),
                bullets=list(obj.get("bullets") or []),
            )

        return cls(
            title=data.get("title", "Результат расчёта"),
            source_data=[row(x) for x in data.get("source_data") or []],
            rate_licenses=[rate_license(x) for x in data.get("rate_licenses") or []],
            contract_rate=data.get("contract_rate"),
            quarter_amount=data.get("quarter_amount"),
            minimum_licenses=[minimum_license(x) for x in data.get("minimum_licenses") or []],
            minimum_rows=[row(x) for x in data.get("minimum_rows") or []],
            contract_terms=[terms(x) for x in data.get("contract_terms") or []],
            comments=list(data.get("comments") or []),
        )
