"""Turn past Ahrens estimates into Ahrens' own rates.

Claude reads each past estimate and measures the building, writing one row per job to past_jobs.csv.
This module does the math: medians by building type and size band, with the jobs each number rests on.
"""
from __future__ import annotations
import csv
import statistics
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill

from .pricing import band

CSV_COLUMNS = ["job_id", "date", "client", "project", "address", "building_type", "gross_sf", "max_floors",
               "lump_sum", "crew_weeks", "crew_size", "scope_included", "won", "actual_cost", "source_file",
               "located"]
MIN_JOBS = 3
WORKDAYS_PER_WEEK = 5


def load_past_jobs(csv_path: Path) -> list[dict]:
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def _num(v) -> float | None:
    try:
        x = float(str(v).replace("$", "").replace(",", "").strip())
        return x if x > 0 else None
    except ValueError:
        return None


def calibrate(jobs: list[dict]) -> dict[str, dict]:
    samples: dict[str, list[tuple[str, float]]] = {}
    for j in jobs:
        if str(j.get("located", "")).strip().upper() != "Y":
            continue
        sf, price, floors = _num(j.get("gross_sf")), _num(j.get("lump_sum")), _num(j.get("max_floors")) or 1
        if not sf:
            continue
        scope = {s.strip() for s in str(j.get("scope_included", "")).split(";")}
        if price and "demolition" in scope and j.get("building_type"):
            samples.setdefault(f"sfrate_{j['building_type'].strip().lower()}", []).append((j["job_id"], price / sf))
        weeks = _num(j.get("crew_weeks"))
        if weeks:
            samples.setdefault(f"sf_per_crew_day_{band(floors)}", []).append(
                (j["job_id"], sf / (weeks * WORKDAYS_PER_WEEK)))
    out = {}
    for key, pts in samples.items():
        if len(pts) < MIN_JOBS:
            continue
        vals = [v for _, v in pts]
        out[key] = {"value": round(statistics.median(vals), 2), "n": len(pts), "jobs": [i for i, _ in pts],
                    "spread": (round(min(vals), 2), round(max(vals), 2))}
    return out


def write_calibration(result: dict, jobs: list[dict], path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "Jobs"
    ws.append(CSV_COLUMNS + ["$/gross SF"])
    for j in jobs:
        sf, price = _num(j.get("gross_sf")), _num(j.get("lump_sum"))
        ws.append([j.get(c, "") for c in CSV_COLUMNS] + [round(price / sf, 2) if sf and price else None])
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True)
    d = wb.create_sheet("Derived Rates")
    d.append(["Rate key", "Value", "Jobs (n)", "Low", "High", "Job ids"])
    for c in d[1]:
        c.font = Font(name="Arial", bold=True)
    for k, v in sorted(result.items()):
        d.append([k, v["value"], v["n"], v["spread"][0], v["spread"][1], ", ".join(v["jobs"])])
    d.append([])
    d.append([f"Rates need at least {MIN_JOBS} located jobs. $/SF uses only jobs whose scope included demolition."])
    wb.save(path)


def apply_to_rate_card(result: dict, rate_card: Path) -> list[str]:
    wb = openpyxl.load_workbook(rate_card); ws = wb["Rates"]
    changed = []
    for r in range(5, ws.max_row + 1):
        key = ws.cell(r, 1).value
        if key in result:
            v = result[key]
            ws.cell(r, 4).value = v["value"]
            ws.cell(r, 4).font = Font(name="Arial")
            ws.cell(r, 4).fill = PatternFill(fill_type=None)
            ws.cell(r, 5).value = f"Ahrens past jobs (n={v['n']}): {', '.join(v['jobs'])}"
            ws.cell(r, 6).value = "N"
            changed.append(key)
    wb.save(rate_card)
    return changed
