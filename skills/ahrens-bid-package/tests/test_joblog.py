import openpyxl
from bidpkg.joblog import append_job, COLUMNS
from bidpkg.pricing import price_crew_time, price_per_sf

def test_append_creates_then_appends(tmp_path, job, rates):
    log = tmp_path / "Bid Log.xlsx"
    a, b = price_crew_time(job, rates), price_per_sf(job, rates)
    assert append_job(log, job, a, b, tmp_path) == 2
    assert append_job(log, job, a, b, tmp_path) == 3
    ws = openpyxl.load_workbook(log).active
    assert [c.value for c in ws[1]] == COLUMNS
    assert ws["F2"].value == a.total and ws["G2"].value == b.total and ws["H2"].value is None
    assert ws["E2"].value == 20000 and ws["C2"].value == "Fixture Bldg"
    assert ws["H2"].fill.fgColor.rgb.endswith("FFFF00")
