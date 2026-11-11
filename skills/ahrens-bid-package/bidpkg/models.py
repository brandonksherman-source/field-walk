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
        from .scope import default_switches
        sw = default_switches()
        sw.update({k: _as_bool(k, v) for k, v in (d.get("switches") or {}).items()})
        d["switches"] = sw
        known = {f.name for f in fields(cls)}
        return cls(**{k: v for k, v in d.items() if k in known})


def _as_bool(key: str, v) -> bool:
    if isinstance(v, bool):
        return v
    t = str(v).strip().lower()
    if t in ("y", "yes", "true", "1", "on", "include", "included"):
        return True
    if t in ("n", "no", "false", "0", "off", "exclude", "excluded"):
        return False
    raise ValueError(f"switch {key!r} must be Y/N (true/false), got {v!r}")


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
    qty_ref: str = ""  # how the workbook derives qty: "q:<quantity key>", "round:<key>" or a constant


@dataclass
class PriceResult:
    method: str
    lines: list[Line]
    direct: float
    markups: dict[str, float]
    total: float
    warning: str | None = None
