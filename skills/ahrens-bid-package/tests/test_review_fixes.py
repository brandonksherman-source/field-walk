"""Regression tests for the final-review findings."""
import json, subprocess, sys
from pathlib import Path
import openpyxl, pytest, docx
from bidpkg.models import Job
from bidpkg.pricing import quantities, price_crew_time, price_per_sf, round_half_up
from bidpkg.workbook import write_workbook, recalc
from bidpkg.letter import render_letter
from bidpkg.joblog import append_job

ROOT = Path(__file__).resolve().parents[1]
FIXTURE_JOB = ROOT / "tests" / "fixtures" / "fixture_job.json"

def run(args):
    r = subprocess.run([sys.executable, "-m", "bidpkg.cli", *map(str, args)], cwd=ROOT, capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    return r.stdout

def summary(wb, label):
    ws = wb["Summary"]
    return next(ws.cell(r, 2).value for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == label)

def text_of(p):
    d = docx.Document(p); return "\n".join(x.text for x in d.paragraphs)

# 1 – a switch flipped (or price entered) in the workbook reaches the letter after `rebuild`
def test_rebuild_reads_switches_and_price_from_workbook(tmp_path):
    run(["init", "--root", tmp_path]); rc = tmp_path / "Ahrens Rate Card.xlsx"
    folder = Path(run(["build", FIXTURE_JOB, "--root", tmp_path, "--rates", rc]).strip().splitlines()[-1])
    wbp = folder / "Fixture Bldg - Takeoff.xlsx"
    wb = openpyxl.load_workbook(wbp); ws = wb["Crew & Time"]
    r = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 8).value == "scrap_credit")
    ws.cell(r, 1).value = "N"; wb.save(wbp)
    run(["rebuild", folder, "--root", tmp_path, "--rates", rc])
    inc, exc = text_of(folder / "Fixture Bldg - Proposal.docx").split("Exclusions:")
    assert "All salvage generated (paid to owner)" in exc and "retained by Ahrens" not in inc
    job = Job.from_dict(json.loads((folder / "job.json").read_text()))
    assert job.switches["scrap_credit"] is False
    from bidpkg.ratecard import load_rates
    assert f"LUMP SUM: ${price_crew_time(job, load_rates(rc)).total:,.2f}" in inc + exc
    # reviewer's "Price to send" wins on the letter
    wb = openpyxl.load_workbook(wbp); wb["Summary"]["B11"] = 123456; wb.save(wbp)
    run(["rebuild", folder, "--root", tmp_path, "--rates", rc])
    assert "LUMP SUM: $123,456.00" in text_of(folder / "Fixture Bldg - Proposal.docx")
    assert openpyxl.load_workbook(tmp_path / "Bid Log.xlsx").active.max_row == 2   # rebuild doesn't add a log row

# 2 – overrides of base quantities flow into gross SF (Python and workbook)
def test_basement_override_flows_into_gross(tmp_path, job, rates):
    job.quantities_override["basement_sf"] = 10000
    assert quantities(job, rates)["gross_sf"] == 30000
    p = tmp_path / "t.xlsx"; write_workbook(job, rates, p); assert recalc(p) == {}
    v = openpyxl.load_workbook(p, data_only=True)
    assert summary(v, "Crew & time price") == price_crew_time(job, rates).total
    assert summary(v, "Total gross SF") == 30000

# 3 – .5 rounds the same way as Excel
def test_round_half_up():
    assert round_half_up(2.5) == 3 and round_half_up(480.5) == 481 and round_half_up(-2.5) == -3

def test_half_fence_matches_workbook(tmp_path, job, rates):
    job.quantities_override["fence_lf"] = 480.5
    p = tmp_path / "t.xlsx"; write_workbook(job, rates, p); recalc(p)
    v = openpyxl.load_workbook(p, data_only=True)
    assert summary(v, "Crew & time price") == price_crew_time(job, rates).total
    assert summary(v, "$/SF price") == price_per_sf(job, rates).total

# 4 – text starting with "=" stays text
def test_equals_text_stays_literal(tmp_path, job, rates):
    job.questions = [("=SUM(1,2)", "=HYPERLINK(\"x\")", "HIGH")]
    job.project = "=Bad Name"
    p = tmp_path / "t.xlsx"; write_workbook(job, rates, p); recalc(p)
    v = openpyxl.load_workbook(p, data_only=True)
    assert v["Bid Checklist"]["A6"].value == "=SUM(1,2)"
    log = tmp_path / "log.xlsx"; append_job(log, job, price_crew_time(job, rates), price_per_sf(job, rates), tmp_path)
    assert openpyxl.load_workbook(log).active["C2"].value == "=Bad Name"
    assert openpyxl.load_workbook(log).active["C2"].data_type == "s"

# 5 – blank recipient
@pytest.mark.parametrize("who", ["", "   "])
def test_blank_recipient_placeholder(tmp_path, job, who):
    job.recipient = who
    d = docx.Document(render_letter(job, 1000, tmp_path)[0])
    runs = [r for p in d.paragraphs for r in p.runs]
    rec = next(r for r in runs if r.text == "[RECIPIENT]"); first = next(r for r in runs if r.text == "[FIRST NAME]")
    assert rec.font.highlight_color is not None and first.font.highlight_color is not None

# 6 – text switch values are understood, junk is refused
def test_text_switch_values(job):
    d = job.to_dict(); d["switches"] = {"scrap_credit": "N", "hazmat": "yes", "backfill": False}
    j = Job.from_dict(d)
    assert j.switches["scrap_credit"] is False and j.switches["hazmat"] is True and j.switches["backfill"] is False
    d["switches"] = {"scrap_credit": "maybe"}
    with pytest.raises(ValueError, match="scrap_credit"):
        Job.from_dict(d)

# 7 – the workbook warns when the reviewer's switches push the price below zero
def test_workbook_warns_when_floored(tmp_path, job, rates):
    for k in job.switches: job.switches[k] = k == "scrap_credit"
    p = tmp_path / "t.xlsx"; write_workbook(job, rates, p); recalc(p)
    v = openpyxl.load_workbook(p, data_only=True)
    assert summary(v, "Price check") == "PRICE FLOORED AT $0 – check scrap credit and switches"

def test_workbook_price_check_blank_when_ok(tmp_path, job, rates):
    p = tmp_path / "t.xlsx"; write_workbook(job, rates, p); recalc(p)
    assert summary(openpyxl.load_workbook(p, data_only=True), "Price check") in ("", None)

# minor fixed alongside #2: the Bid Log uses the same gross SF as the workbook
def test_bid_log_gross_uses_override(tmp_path, job, rates):
    job.quantities_override["gross_sf"] = 50000
    log = tmp_path / "log.xlsx"; append_job(log, job, price_crew_time(job, rates), price_per_sf(job, rates), tmp_path)
    assert openpyxl.load_workbook(log).active["E2"].value == 50000
