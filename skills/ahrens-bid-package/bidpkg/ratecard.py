"""The Ahrens rate card: one sheet of rates the skill reads on every run."""
from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

PLACEHOLDER_SOURCE = "market placeholder – replace with Ahrens rates/quotes"
BLUE = Font(name="Arial", color="0000FF")
BLACK = Font(name="Arial")
BOLD = Font(name="Arial", bold=True)
YELLOW = PatternFill("solid", fgColor="FFFF00")
NOFILL = PatternFill(fill_type=None)


class RateCardError(ValueError):
    pass


@dataclass
class Rate:
    value: float
    unit: str
    source: str
    placeholder: bool


# (section, key, item, unit, value)
SEED = [
    ("General", "mobilization", "Mobilization", "LS", 6500),
    ("General", "highreach_mob", "High-reach mobilization (9+ floors)", "LS", 45000),
    ("General", "permit_per_bldg", "Demo permit", "per bldg", 350),
    ("General", "fencing_lf", "Construction fence & protection", "LF", 6.00),
    ("General", "erosion_ea", "Erosion control / inlet protection", "EA", 2500),
    ("General", "swppp_ls", "SWPPP measures during demolition", "LS", 2500),
    ("Utilities", "electric_disconnect", "Electric disconnect", "EA", 350),
    ("Utilities", "gas_cut_cap", "Gas cut & cap at main", "EA", 2500),
    ("Utilities", "sewer_cap", "Sewer lateral cap", "EA", 2800),
    ("Utilities", "water_kill", "Water kill at main", "EA", 2000),
    ("Utilities", "tap_destroy", "Tap destroy", "EA", 1500),
    ("Hazmat", "hazmat_survey_per_bldg", "Asbestos/lead survey (NESHAP)", "per bldg", 1200),
    ("Hazmat", "abatement_allowance_per_bldg", "Abatement allowance (pre-1980 bldg)", "per bldg", 8000),
    ("Hazmat", "universal_waste_sf", "Universal waste (bulbs, ballasts, Freon)", "SF", 0.35),
    ("Demolition", "mech_demo_sf", "Mechanical demo (reference)", "SF", 3.50),
    ("Demolition", "highrise_premium_sf", "High-rise premium (reference)", "SF", 4.00),
    ("Demolition", "debris_t_per_sf", "Debris tons per gross SF", "t/SF", 0.05),
    ("Demolition", "steel_t_per_sf", "Steel scrap tons per gross SF", "t/SF", 0.0055),
    ("Haul & disposal", "haul_load", "Haul", "$/load", 375),
    ("Haul & disposal", "tons_per_load", "Tons per load", "t/load", 20),
    ("Haul & disposal", "disposal_ton", "C&D disposal / tipping", "$/ton", 35),
    ("Haul & disposal", "scrap_credit_ton", "Scrap steel credit", "$/ton", 180),
    ("Site", "fill_lcy", "Clean fill delivered", "$/loose CY", 18),
    ("Site", "place_compact_cy", "Place & compact", "$/CY", 8),
    ("Site", "seed_straw_sf", "Seed / straw", "$/SF", 0.10),
    ("Site", "tree_ea", "Tree & stump removal", "EA", 900),
    ("Crew day rates", "operator_day", "Operator", "$/day", 720),
    ("Crew day rates", "laborer_day", "Laborer", "$/day", 480),
    ("Crew day rates", "foreman_day", "Foreman", "$/day", 800),
    ("Crew day rates", "driver_day", "Driver", "$/day", 560),
    ("Equipment day rates", "excavator_day", "Excavator (mid)", "$/day", 1400),
    ("Equipment day rates", "excavator_large_day", "Excavator (large)", "$/day", 2600),
    ("Equipment day rates", "highreach_day", "High-reach excavator", "$/day", 4500),
    ("Equipment day rates", "skidsteer_day", "Skid steer", "$/day", 450),
    ("Equipment day rates", "loader_day", "Loader", "$/day", 900),
    ("Production", "sf_per_crew_day_low", "Gross SF per crew-day, 1–3 floors", "SF/day", 6000),
    ("Production", "sf_per_crew_day_mid", "Gross SF per crew-day, 4–8 floors", "SF/day", 3500),
    ("Production", "sf_per_crew_day_high", "Gross SF per crew-day, 9+ floors", "SF/day", 1500),
    ("$/SF by building type", "sfrate_residential", "Residential / wood frame (demo only)", "$/SF", 3.00),
    ("$/SF by building type", "sfrate_masonry", "Masonry / brick (demo only)", "$/SF", 3.50),
    ("$/SF by building type", "sfrate_steel", "Steel frame (demo only)", "$/SF", 4.00),
    ("$/SF by building type", "sfrate_concrete", "Concrete frame (demo only)", "$/SF", 4.50),
    ("$/SF by building type", "sfrate_highrise", "High-rise 9+ floors (demo only)", "$/SF", 7.50),
    ("Markups", "contingency_pct", "Contingency", "%", 0.10),
    ("Markups", "overhead_pct", "Overhead", "%", 0.10),
    ("Markups", "profit_pct", "Profit", "%", 0.10),
    ("Markups", "bond_pct", "Bond (when required)", "%", 0.015),
]


def create_rate_card(path: Path) -> None:
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Rates"
    ws["A1"] = "Ahrens Contracting – Rate Card"; ws["A1"].font = Font(name="Arial", bold=True, size=14)
    ws["A2"] = ("Legend: blue on yellow = placeholder – paste Ahrens numbers over it and change Placeholder? to N · "
                "black = Ahrens rate. The bid skill reads column D by the key in column A.")
    ws["A2"].font = BLACK
    for i, h in enumerate(["Key", "Item", "Unit", "Rate", "Source", "Placeholder?", "Section"], 1):
        c = ws.cell(4, i, h); c.font = BOLD
    for r, (section, key, item, unit, value) in enumerate(SEED, 5):
        vals = [key, item, unit, value, PLACEHOLDER_SOURCE, "Y", section]
        for i, v in enumerate(vals, 1):
            c = ws.cell(r, i, v); c.font = BLACK
        ws.cell(r, 4).font = BLUE; ws.cell(r, 4).fill = YELLOW
        ws.cell(r, 4).number_format = "0.0%" if unit == "%" else "#,##0.00##"
    for col, w in zip("ABCDEFG", (28, 40, 12, 12, 55, 13, 22)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A5"
    wb.save(path)


def load_rates(path: Path) -> dict[str, Rate]:
    ws = openpyxl.load_workbook(path, data_only=True)["Rates"]
    out: dict[str, Rate] = {}
    for r in range(5, ws.max_row + 1):
        key = ws.cell(r, 1).value
        if not key:
            continue
        v = ws.cell(r, 4).value
        if v is None or isinstance(v, bool) or not isinstance(v, (int, float)):
            raise RateCardError(f"Rates!D{r} ({key}) is blank or not a number: {v!r}")
        out[str(key).strip()] = Rate(float(v), str(ws.cell(r, 3).value or ""), str(ws.cell(r, 5).value or ""),
                                     str(ws.cell(r, 6).value or "Y").strip().upper() != "N")
    return out
