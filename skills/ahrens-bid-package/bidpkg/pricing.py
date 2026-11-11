"""Quantities and the two pricing methods (crew & time, $/SF check).

Every line is qty × rate so the workbook can show the same math as formulas.
"""
from __future__ import annotations
import math
from .models import Job, Line, PriceResult
from .scope import SCOPE_LINES

HIGHRISE_FLOORS = 9
MARKUP_ORDER = ("contingency_pct", "overhead_pct", "profit_pct")


def band(max_floors: float) -> str:
    if max_floors >= HIGHRISE_FLOORS:
        return "high"
    return "mid" if max_floors >= 4 else "low"


def quantities(job: Job, rates) -> dict[str, float]:
    S = job.structures
    q: dict[str, float] = {
        "above_grade_sf": sum(s.above_grade_sf for s in S),
        "basement_sf": sum(s.basement_sf for s in S),
        "footprint_sf": sum(s.footprint_sf for s in S),
        "buildings": float(len(S)),
        "max_floors": max((s.floors for s in S), default=0),
        "trees": 0.0,
        "pre1980_bldgs": float(sum(1 for s in S if _year(s.year_built) is None or _year(s.year_built) < 1980)),
    }
    q["gross_sf"] = q["above_grade_sf"] + q["basement_sf"]
    q.update({k: v for k, v in job.quantities_override.items() if k in q})
    q["debris_tons"] = q["gross_sf"] * rates["debris_t_per_sf"].value
    q["steel_tons"] = q["gross_sf"] * rates["steel_t_per_sf"].value
    q["haul_loads"] = float(math.ceil(round(q["debris_tons"] / rates["tons_per_load"].value, 6)))
    q["backfill_cy"] = q["basement_sf"] * 12 / 27 * 1.25
    q["fence_lf"] = 4 * math.sqrt(q["footprint_sf"]) * 1.2
    q["seed_sf"] = q["footprint_sf"]
    q["crew_days"] = float(math.ceil(round(q["gross_sf"] / rates[f"sf_per_crew_day_{band(q['max_floors'])}"].value, 6)))
    q["equipment_days"] = q["crew_days"]
    q["disposal_tons"] = q["debris_tons"] - q["steel_tons"]
    q.update(job.quantities_override)
    return q


def _r(rates, *keys) -> tuple[float, str]:
    total = sum(rates[k].value for k in keys)
    return total, " + ".join(keys) + (" (placeholder)" if any(rates[k].placeholder for k in keys) else "")


def _common_lines(job: Job, rates, q) -> list[tuple[str, str, float, str, float, str]]:
    """(key, label, qty, unit, rate, rate_source) for every line except demolition."""
    hi = q["max_floors"] >= HIGHRISE_FLOORS
    mob = _r(rates, "mobilization", "highreach_mob") if hi else _r(rates, "mobilization")
    util = _r(rates, "electric_disconnect", "gas_cut_cap", "sewer_cap", "water_kill")
    out = [
        ("mobilization", "Mobilization", 1, "LS", *mob, "1"),
        ("permits", "Demolition permits", q["buildings"], "EA", *_r(rates, "permit_per_bldg"), "q:buildings"),
        ("fencing", "Construction fence & protection", round(q["fence_lf"]), "LF", *_r(rates, "fencing_lf"), "round:fence_lf"),
        ("utilities", "Utility disconnects / cut & cap (E, G, W, S)", q["buildings"], "bldg", *util, "q:buildings"),
        ("tap_destroys", "Tap destroys", q["buildings"], "EA", *_r(rates, "tap_destroy"), "q:buildings"),
        ("hazmat", "Hazmat survey (NESHAP)", q["buildings"], "bldg", *_r(rates, "hazmat_survey_per_bldg"), "q:buildings"),
        ("hazmat", "Abatement allowance (pre-1980 or unknown year)", q["pre1980_bldgs"], "bldg",
         *_r(rates, "abatement_allowance_per_bldg"), "q:pre1980_bldgs"),
        ("hazmat", "Universal waste (bulbs, ballasts, Freon)", q["above_grade_sf"], "SF", *_r(rates, "universal_waste_sf"), "q:above_grade_sf"),
        ("haul", "Haul debris", q["haul_loads"], "load", *_r(rates, "haul_load"), "q:haul_loads"),
        ("disposal", "Tipping / disposal (debris less steel)", q["disposal_tons"], "ton",
         *_r(rates, "disposal_ton"), "q:disposal_tons"),
        ("scrap_credit", "Scrap steel credit", q["steel_tons"], "ton",
         -rates["scrap_credit_ton"].value, _r(rates, "scrap_credit_ton")[1] + " – credit", "q:steel_tons"),
        ("backfill", "Backfill basements (loose CY)", q["backfill_cy"], "CY", *_r(rates, "fill_lcy", "place_compact_cy"), "q:backfill_cy"),
        ("seed_straw", "Seed / straw", q["seed_sf"], "SF", *_r(rates, "seed_straw_sf"), "q:seed_sf"),
        ("trees", "Tree & stump removal", q["trees"], "EA", *_r(rates, "tree_ea"), "q:trees"),
        ("swppp", "Erosion control / SWPPP", 1, "LS", *_r(rates, "swppp_ls", "erosion_ea"), "1"),
    ]
    return out


def _year(y) -> int | None:
    try:
        return int(str(y).strip()[:4])
    except (ValueError, TypeError):
        return None


def _finish(method: str, job: Job, rates, raw) -> PriceResult:
    lines = []
    for key, label, qty, unit, rate, src, ref in raw:
        inc = job.switches.get(key, True)
        lines.append(Line(key, label, float(qty), unit, float(rate), float(qty) * float(rate) if inc else 0.0, inc, src, ref))
    direct = sum(l.extended for l in lines)
    running, markups = direct, {}
    for k in MARKUP_ORDER:
        amt = running * rates[k].value
        markups[k] = amt
        running += amt
    total, warning = float(round(running)), None
    if total < 0:
        total, warning = 0.0, "Price below zero – check scrap credit and switches"
    return PriceResult(method, lines, direct, markups, total, warning)


def demolition_lines_crew(job: Job, rates, q):
    hi = q["max_floors"] >= HIGHRISE_FLOORS
    eq = "highreach_day" if hi else ("excavator_large_day" if band(q["max_floors"]) == "mid" else "excavator_day")
    return [
        ("demolition", "Demolition crew (operator + laborer + foreman)", q["crew_days"], "crew-day",
         *_r(rates, "operator_day", "laborer_day", "foreman_day"), "q:crew_days"),
        ("demolition", "Demolition equipment", q["equipment_days"], "day", *_r(rates, eq), "q:equipment_days"),
    ]


def price_crew_time(job: Job, rates) -> PriceResult:
    q = quantities(job, rates)
    return _finish("crew_time", job, rates, demolition_lines_crew(job, rates, q) + _common_lines(job, rates, q))


def price_per_sf(job: Job, rates) -> PriceResult:
    q = quantities(job, rates)
    key = f"sfrate_{job.building_type}"
    if key not in rates:
        raise KeyError(f"No $/SF rate for building type {job.building_type!r} (expected {key} in the rate card)")
    demo = [("demolition", f"Demolition – {job.building_type} ($/gross SF)", q["gross_sf"], "SF", *_r(rates, key), "q:gross_sf")]
    return _finish("per_sf", job, rates, demo + _common_lines(job, rates, q))
