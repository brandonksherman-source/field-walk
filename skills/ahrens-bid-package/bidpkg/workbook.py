"""The internal takeoff workbook: same tabs, same look, formulas throughout.

One Include? cell per scope key (on Crew & Time) drives both prices; every other line of that key and the
$-SF Check sheet link to it.
"""
from __future__ import annotations
import shutil
import subprocess
import tempfile
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from .models import Job, PriceResult
from .pricing import quantities, price_crew_time, price_per_sf, band, MARKUP_ORDER

TABS = ["Summary", "SF Takeoff", "Quantities", "Crew & Time", "$-SF Check", "Bid Checklist", "Sources"]
LEGEND = "Legend: blue = input · black = formula · green = link from another sheet · yellow fill = verify / placeholder"

F_BLACK = Font(name="Arial")
F_BLUE = Font(name="Arial", color="0000FF")
F_GREEN = Font(name="Arial", color="008000")
F_BOLD = Font(name="Arial", bold=True)
F_TITLE = Font(name="Arial", bold=True, size=13)
YELLOW = PatternFill("solid", fgColor="FFFF00")
MONEY = '"$"#,##0;[Red]("$"#,##0)'
MONEY2 = '"$"#,##0.00;[Red]("$"#,##0.00)'
NUM = "#,##0"
NUM2 = "#,##0.00##"


def _q(name: str) -> str:
    return "'" + name.replace("'", "''") + "'"


def _header(ws, tab: str, job: Job, legend: bool = True) -> None:
    ws["A1"] = f"{tab} – {job.project} ({job.address})"
    ws["A1"].font = F_TITLE
    due = f" · Bid due: {job.due_date}" if job.due_date else " · Bid due: NOT GIVEN – confirm"
    who = f"Requested by: {job.recipient or '[RECIPIENT]'}, {job.client}" if job.client else "Requested by: [CLIENT]"
    ws["A2"] = f"{who}{due} · Prepared {job.date:%m/%d/%Y} by Ahrens Contracting (DRAFT – for review)"
    ws["A2"].font = F_BLACK
    if legend:
        ws["A3"] = LEGEND
        ws["A3"].font = F_BLACK


def _row(ws, r: int, values, fonts=None, formats=None) -> None:
    for i, v in enumerate(values, 1):
        if v is None:
            continue
        c = ws.cell(r, i, v)
        c.font = (fonts or {}).get(i, F_BLACK)
        if formats and i in formats:
            c.number_format = formats[i]


def _widths(ws, widths) -> None:
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _sf_takeoff(wb, job: Job) -> dict:
    ws = wb.create_sheet("SF Takeoff")
    _header(ws, "SF Takeoff", job)
    head = ["Structure", "Address / Parcel", "Footprint SF (±10%)", "Floors", "Published GSF (if any)",
            "Above-grade gross SF", "Basement SF", "Total gross SF", "Year built", "Owner of record", "Demo note",
            "Source"]
    _row(ws, 5, head, {i: F_BOLD for i in range(1, 13)})
    first = 6
    r = first
    for s in job.structures:
        _row(ws, r, [s.name, f"{s.address} · {s.parcel}".strip(" ·"), s.footprint_sf, s.floors,
                     s.published_gsf if s.published_gsf else None,
                     f'=IF(E{r}<>"",E{r},C{r}*D{r})', s.basement_sf, f"=F{r}+G{r}", s.year_built, s.owner, s.note,
                     s.source],
             {3: F_BLUE, 4: F_BLUE, 5: F_BLUE, 7: F_BLUE},
             {3: NUM, 4: "0.0", 5: NUM, 6: NUM, 7: NUM, 8: NUM})
        ws.cell(r, 3).fill = YELLOW
        ws.cell(r, 7).fill = YELLOW
        r += 1
    last = max(r - 1, first)
    tot = last + 1
    _row(ws, tot, ["TOTAL – in scope", None, f"=SUM(C{first}:C{last})", None, None, f"=SUM(F{first}:F{last})",
                   f"=SUM(G{first}:G{last})", f"=SUM(H{first}:H{last})"],
         {1: F_BOLD}, {3: NUM, 6: NUM, 7: NUM, 8: NUM})
    rr = tot + 2
    ws.cell(rr, 1, "STRUCTURES TO REMAIN / PROTECT (not in quote)").font = F_BOLD
    for item in job.to_remain:
        rr += 1
        ws.cell(rr, 1, item).font = F_BLACK
    _widths(ws, [30, 34, 14, 8, 14, 16, 13, 15, 11, 30, 50, 40])
    ws.freeze_panes = "A6"
    return {"first": first, "last": last, "tot": tot}


# key: (label, unit, kind, input_value_or_None, basis)
def _quantity_rows(job, rates, q, sf):
    t, f, l = sf["tot"], sf["first"], sf["last"]
    S = _q("SF Takeoff")
    b = band(q["max_floors"])
    return [
        ("gross_sf", "Total gross SF (incl. basements)", "SF", f"={S}!H{t}", None, "Link"),
        ("above_grade_sf", "Above-grade gross SF", "SF", f"={S}!F{t}", None, "Link"),
        ("basement_sf", "Basement SF", "SF", f"={S}!G{t}", None, "Link"),
        ("footprint_sf", "Footprint total", "SF", f"={S}!C{t}", None, "Link"),
        ("buildings", "Buildings in scope", "EA", f"=COUNT({S}!C{f}:C{l})", None, "Count of structures"),
        ("max_floors", "Tallest structure (floors)", "fl", f"=MAX({S}!D{f}:D{l})", None, "Max of floors"),
        ("pre1980_bldgs", "Buildings pre-1980 or unknown year", "EA", q["pre1980_bldgs"], None,
         "Abatement allowance applies – VERIFY with survey"),
        ("debris_tons", "Total demolition debris", "ton", "=B{gross_sf}*D{r}", rates["debris_t_per_sf"].value,
         "Gross SF × t/SF – ASSUMPTION"),
        ("steel_tons", "Structural / rebar steel (scrap)", "ton", "=B{gross_sf}*D{r}", rates["steel_t_per_sf"].value,
         "Gross SF × t/SF"),
        ("disposal_tons", "Debris to landfill (debris – steel)", "ton", "=B{debris_tons}-B{steel_tons}", None,
         "Steel goes to scrap"),
        ("haul_loads", "Haul loads", "load", "=ROUNDUP(B{debris_tons}/D{r},0)", rates["tons_per_load"].value,
         "Tons ÷ t/load"),
        ("backfill_cy", "Basement backfill (loose CY)", "CY", "=B{basement_sf}*D{r}/27*1.25", 12,
         "Basement SF × depth ft ÷ 27 × 1.25 swell – VERIFY depth"),
        ("fence_lf", "Construction fence", "LF", "=4*SQRT(B{footprint_sf})*1.2", None,
         "Square perimeter of footprint × 1.2"),
        ("seed_sf", "Seed / straw area", "SF", "=B{footprint_sf}", None, "Footprint"),
        ("trees", "Trees to remove", "EA", q["trees"], None, "From site photos / aerial – VERIFY"),
        ("crew_days", "Demolition crew-days", "day", "=ROUNDUP(B{gross_sf}/D{r},0)",
         rates[f"sf_per_crew_day_{b}"].value, f"Gross SF ÷ SF per crew-day ({b} band) – PLACEHOLDER production"),
        ("equipment_days", "Equipment days", "day", "=B{crew_days}", None, "Same as crew-days"),
    ]


def _quantities(wb, job: Job, rates, q, sf) -> dict[str, int]:
    ws = wb.create_sheet("Quantities")
    _header(ws, "Quantities", job)
    _row(ws, 5, ["Item", "Qty", "Unit", "Input", "Basis / assumption", "Key"], {i: F_BOLD for i in range(1, 7)})
    spec = _quantity_rows(job, rates, q, sf)
    rows = {key: 6 + i for i, (key, *_r) in enumerate(spec)}
    for key, label, unit, formula, inp, basis in spec:
        r = rows[key]
        if key in job.quantities_override:
            val, font, basis = job.quantities_override[key], F_BLUE, basis + " · OVERRIDE"
        elif isinstance(formula, str):
            val = formula.format(r=r, **rows)
            font = F_GREEN if "!" in val else F_BLACK
        else:
            val, font = formula, F_BLUE
        _row(ws, r, [label, val, unit, inp, basis, key], {2: font, 4: F_BLUE}, {2: NUM2, 4: NUM2})
        if inp is not None or font is F_BLUE:
            ws.cell(r, 4 if inp is not None else 2).fill = YELLOW
    ws.column_dimensions["F"].hidden = True
    _widths(ws, [40, 14, 8, 10, 60, 16])
    return rows


def _qty_formula(ref: str, qrows: dict[str, int]) -> tuple[object, Font]:
    Q = _q("Quantities")
    if ref.startswith("q:"):
        return f"={Q}!B{qrows[ref[2:]]}", F_GREEN
    if ref.startswith("round:"):
        return f"=ROUND({Q}!B{qrows[ref[6:]]},0)", F_GREEN
    return float(ref), F_BLUE


def _pricing_sheet(wb, tab: str, job: Job, res: PriceResult, rates, qrows, include_cells: dict[str, str] | None):
    """Writes a pricing sheet. include_cells=None → this sheet owns the switches; returns key → cell address."""
    ws = wb.create_sheet(tab)
    _header(ws, tab, job)
    ws["A4"] = "Include? Y/N on Crew & Time switches a scope line for BOTH prices and moves it to the letter's Exclusions."
    ws["A4"].font = F_BLACK
    _row(ws, 5, ["Include?", "Line", "Qty", "Unit", "Unit rate ($)", "Extended ($)", "Rate source", "Key"],
         {i: F_BOLD for i in range(1, 9)})
    dv = DataValidation(type="list", formula1='"Y,N"', allow_blank=False)
    ws.add_data_validation(dv)
    owned: dict[str, str] = {}
    CT = _q("Crew & Time")
    r = 6
    for line in res.lines:
        if include_cells is None and line.key not in owned:
            ws.cell(r, 1, "Y" if line.included else "N").font = F_BLUE
            dv.add(f"A{r}")
            owned[line.key] = f"A{r}"
        elif include_cells is None:
            ws.cell(r, 1, f"={owned[line.key]}").font = F_BLACK
        else:
            ws.cell(r, 1, f"={CT}!{include_cells[line.key]}").font = F_GREEN
        qty, qfont = _qty_formula(line.qty_ref or str(line.qty), qrows)
        _row(ws, r, [None, line.label, qty, line.unit, line.rate, f'=IF(A{r}="Y",C{r}*E{r},0)', line.rate_source,
                     line.key], {3: qfont, 5: F_BLUE}, {3: NUM2, 5: MONEY2, 6: MONEY})
        if "placeholder" in line.rate_source:
            ws.cell(r, 5).fill = YELLOW
        r += 1
    last = r - 1
    direct = r + 1
    _row(ws, direct, [None, "Direct cost", None, None, None, f"=SUM(F6:F{last})"], {2: F_BOLD}, {6: MONEY})
    prev = f"F{direct}"
    running_terms = [prev]
    rr = direct
    for k in MARKUP_ORDER:
        rr += 1
        _row(ws, rr, [None, rates_label(k), None, "%", rates[k].value, f"=({'+'.join(running_terms)})*E{rr}",
                      rates[k].source], {5: F_BLUE}, {5: "0.0%", 6: MONEY})
        if rates[k].placeholder:
            ws.cell(rr, 5).fill = YELLOW
        running_terms.append(f"F{rr}")
    total = rr + 1
    _row(ws, total, [None, "BASE BID (lump sum)", None, None, None,
                     f"=MAX(0,ROUND({'+'.join(running_terms)},0))"], {2: F_BOLD, 6: F_BOLD}, {6: MONEY})
    _row(ws, total + 1, [None, "Base bid with bond (if required)", None, "%", rates["bond_pct"].value,
                         f"=ROUND(F{total}*(1+E{total + 1}),0)"], {5: F_BLUE}, {5: "0.0%", 6: MONEY})
    ws.column_dimensions["H"].hidden = True
    _widths(ws, [9, 50, 13, 9, 14, 15, 55, 14])
    ws.freeze_panes = "A6"
    return owned, total


def rates_label(k: str) -> str:
    return {"contingency_pct": "Contingency", "overhead_pct": "Overhead", "profit_pct": "Profit"}[k]


def _checklist(wb, job: Job) -> None:
    ws = wb.create_sheet("Bid Checklist")
    _header(ws, "Bid Checklist / Questions for client", job, legend=False)
    ws["A3"] = "Status: pick Yes / No / N/A. Open items = No or blank."
    _row(ws, 5, ["Item", "Detail", "Priority", "Status"], {i: F_BOLD for i in range(1, 5)})
    dv = DataValidation(type="list", formula1='"Yes,No,N/A"', allow_blank=True)
    ws.add_data_validation(dv)
    r = 6
    for item, detail, prio in job.questions:
        _row(ws, r, [item, detail, prio, "No"], {4: F_BLUE})
        dv.add(f"D{r}")
        r += 1
    _widths(ws, [44, 90, 10, 10])


def _sources(wb, job: Job, rates) -> None:
    ws = wb.create_sheet("Sources")
    _header(ws, "Sources", job, legend=False)
    r = 5
    ws.cell(r, 1, "Sources").font = F_BOLD
    for s in job.sources:
        r += 1
        ws.cell(r, 1, s).font = F_BLACK
    r += 2
    ws.cell(r, 1, "Placeholder rates used (replace in Ahrens Rate Card.xlsx)").font = F_BOLD
    for k, v in rates.items():
        if v.placeholder:
            r += 1
            ws.cell(r, 1, f"{k} = {v.value:g} {v.unit}").font = F_BLACK
    ws.column_dimensions["A"].width = 120


def _summary(wb, job: Job, ct_total: int, sf_total: int, crew: PriceResult, persf: PriceResult) -> None:
    ws = wb["Summary"]
    _header(ws, "Summary", job)
    CT, SF, SFT, BC = _q("Crew & Time"), _q("$-SF Check"), _q("SF Takeoff"), _q("Bid Checklist")
    rows = [
        ("Crew & time price", f"={CT}!F{ct_total}", F_GREEN, MONEY),
        ("$/SF price", f"={SF}!F{sf_total}", F_GREEN, MONEY),
        ("Difference (crew & time – $/SF)", "=B6-B7", F_BLACK, MONEY),
        ("Total gross SF", f"=Quantities!B6", F_GREEN, NUM),
        ("$ per gross SF (crew & time)", "=IF(B9>0,B6/B9,0)", F_BLACK, MONEY2),
        ("Price to send", None, F_BLUE, MONEY),
        ("Open questions", f'=COUNTIF({BC}!D:D,"No")' + (f'+COUNTBLANK({BC}!D6:D{5 + len(job.questions)})'
                                                          if job.questions else ""), F_GREEN, NUM),
    ]
    for i, (label, val, font, fmt) in enumerate(rows, 6):
        ws.cell(i, 1, label).font = F_BOLD
        c = ws.cell(i, 2, val)
        c.font, c.number_format = font, fmt
    ws["B11"].fill = YELLOW
    ws["C11"] = "Reviewer: enter the number that goes on the letter"
    r = 14
    warnings = [w for w in (crew.warning, persf.warning) if w]
    if warnings:
        ws.cell(r, 1, "WARNINGS").font = F_BOLD
        for w in warnings:
            r += 1
            ws.cell(r, 1, w).font = Font(name="Arial", color="C00000", bold=True)
        r += 2
    ws.cell(r, 1, "Top risks / gaps").font = F_BOLD
    for risk in job.risks:
        r += 1
        ws.cell(r, 1, f"• {risk}").font = F_BLACK
    _widths(ws, [40, 18, 50])


def write_workbook(job: Job, rates, path: Path) -> None:
    q = quantities(job, rates)
    crew, persf = price_crew_time(job, rates), price_per_sf(job, rates)
    wb = openpyxl.Workbook()
    wb.active.title = "Summary"
    sf = _sf_takeoff(wb, job)
    qrows = _quantities(wb, job, rates, q, sf)
    owned, ct_total = _pricing_sheet(wb, "Crew & Time", job, crew, rates, qrows, None)
    _, sf_total = _pricing_sheet(wb, "$-SF Check", job, persf, rates, qrows, owned)
    _checklist(wb, job)
    _sources(wb, job, rates)
    _summary(wb, job, ct_total, sf_total, crew, persf)
    wb.save(path)


def recalc(path: Path) -> dict[str, list[str]]:
    """Recalculate with LibreOffice so the file carries computed values; returns error cells per sheet."""
    path = Path(path)
    with tempfile.TemporaryDirectory() as td:
        profile = Path(td) / "profile"
        out = Path(td) / "out"
        subprocess.run(["soffice", f"-env:UserInstallation=file://{profile}", "--headless", "--calc",
                        "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(out), str(path)],
                       check=True, capture_output=True, timeout=180)
        shutil.copyfile(out / path.name, path)
    errors: dict[str, list[str]] = {}
    for ws in openpyxl.load_workbook(path, data_only=True).worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("Err:"):
                    errors.setdefault(ws.title, []).append(c.coordinate)
                elif isinstance(c.value, str) and c.value.startswith("#") and c.value.rstrip("!?0/").upper() in (
                        "#REF", "#NAME", "#VALUE", "#DIV", "#N/A", "#NUM", "#NULL", "#ERR"):
                    errors.setdefault(ws.title, []).append(c.coordinate)
    return errors
