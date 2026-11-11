"""Bid Log.xlsx: one row per job, so estimates can be compared with what was sent and what it cost."""
from __future__ import annotations
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill
from .models import Job, PriceResult
from .pricing import base_quantities

COLUMNS = ["Date", "Client", "Project", "Address", "Gross SF", "Crew & time price", "$/SF price", "Price sent",
           "Won/Lost", "Actual cost", "Folder"]
YELLOW = PatternFill("solid", fgColor="FFFF00")


def create_log(log_path: Path) -> None:
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Bids"
    ws.append(COLUMNS)
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True)
    for col, w in zip("ABCDEFGHIJK", (12, 22, 34, 40, 12, 16, 14, 14, 10, 14, 60)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    wb.save(log_path)


def append_job(log_path: Path, job: Job, crew: PriceResult, per_sf: PriceResult, folder: Path) -> int:
    log_path = Path(log_path)
    if not log_path.exists():
        create_log(log_path)
    wb = openpyxl.load_workbook(log_path); ws = wb.active
    r = ws.max_row + 1
    gross = base_quantities(job)["gross_sf"]
    vals = [job.date, job.client, job.project, job.address, gross, crew.total, per_sf.total, None, None, None,
            str(folder)]
    for i, v in enumerate(vals, 1):
        c = ws.cell(r, i, v)
        if isinstance(v, str):
            c.data_type = "s"   # never let a name starting with "=" become a formula
        c.font = Font(name="Arial")
    ws.cell(r, 1).number_format = "yyyy-mm-dd"
    for i in (5,):
        ws.cell(r, i).number_format = "#,##0"
    for i in (6, 7, 8, 10):
        ws.cell(r, i).number_format = '"$"#,##0'
    for i in (8, 9, 10):
        ws.cell(r, i).fill = YELLOW
    wb.save(log_path)
    return r
