"""Data shared by every bidpkg module."""
from __future__ import annotations
from dataclasses import dataclass, field, asdict, fields
from datetime import date as _date


@dataclass
class Structure:
    name: str
    address: str
    parcel: str
    footprint_sf: float
    floors: float
    basement_sf: float
    year_built: str
    owner: str
    note: str
    source: str
    published_gsf: float | None = None

    @property
    def above_grade_sf(self) -> float:
        if self.published_gsf:
            return self.published_gsf
        return self.footprint_sf * self.floors

    @property
    def gross_sf(self) -> float:
        return self.above_grade_sf + self.basement_sf


@dataclass
class Job:
    project: str
    client: str
    recipient: str | None
    address: str
    date: _date
    building_type: str = "masonry"
    structures: list[Structure] = field(default_factory=list)
    to_remain: list[str] = field(default_factory=list)
    switches: dict[str, bool] = field(default_factory=dict)
    quantities_override: dict[str, float] = field(default_factory=dict)
    questions: list[tuple[str, str, str]] = field(default_factory=list)
    sources: list[str] = field(default_factory=list)
    included_bullets_extra: list[str] = field(default_factory=list)
    due_date: str | None = None
    risks: list[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        d = asdict(self)
        d["date"] = self.date.isoformat()
        return d

    @classmethod
    def from_dict(cls, d: dict) -> "Job":
        d = dict(d)
        if isinstance(d.get("date"), str):
            d["date"] = _date.fromisoformat(d["date"])
        d["structures"] = [s if isinstance(s, Structure) else Structure(**s) for s in d.get("structures", [])]
        d["questions"] = [tuple(q) for q in d.get("questions", [])]
        if not d.get("switches"):
            from .scope import default_switches
            d["switches"] = default_switches()
        known = {f.name for f in fields(cls)}
        return cls(**{k: v for k, v in d.items() if k in known})


@dataclass
class Line:
    key: str
    label: str
    qty: float
    unit: str
    rate: float
    extended: float
    included: bool
    rate_source: str


@dataclass
class PriceResult:
    method: str
    lines: list[Line]
    direct: float
    markups: dict[str, float]
    total: float
    warning: str | None = None
