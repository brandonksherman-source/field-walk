import openpyxl, pytest
from bidpkg.workbook import write_workbook, recalc, TABS
from bidpkg.pricing import price_crew_time, price_per_sf

def summary_value(wb, label):
    ws = wb["Summary"]
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == label:
            return ws.cell(r, 2).value
    raise KeyError(label)

def set_include(path, sheet, key, value):
    wb = openpyxl.load_workbook(path); ws = wb[sheet]
    r = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 8).value == key)
    ws.cell(r, 1).value = value; wb.save(path)

@pytest.fixture
def built(tmp_path, job, rates):
    p = tmp_path / "t.xlsx"; write_workbook(job, rates, p); return p

def test_tabs_and_clean_recalc(built):
    assert openpyxl.load_workbook(built).sheetnames == ["Summary","SF Takeoff","Quantities","Crew & Time","$-SF Check","Bid Checklist","Sources"] == TABS
    assert recalc(built) == {}

def test_workbook_totals_match_python(built, job, rates):
    recalc(built)
    v = openpyxl.load_workbook(built, data_only=True)
    assert summary_value(v, "Crew & time price") == price_crew_time(job, rates).total
    assert summary_value(v, "$/SF price") == price_per_sf(job, rates).total

def test_flipping_include_cell_changes_both_totals(built, job, rates):
    set_include(built, "Crew & Time", "disposal", "N"); recalc(built)
    job.switches["disposal"] = False
    v = openpyxl.load_workbook(built, data_only=True)
    assert summary_value(v, "Crew & time price") == price_crew_time(job, rates).total
    assert summary_value(v, "$/SF price") == price_per_sf(job, rates).total

def test_switch_off_in_job_writes_N(tmp_path, job, rates):
    job.switches["hazmat"] = False
    p = tmp_path / "t.xlsx"; write_workbook(job, rates, p); recalc(p)
    assert summary_value(openpyxl.load_workbook(p, data_only=True), "Crew & time price") == price_crew_time(job, rates).total

def test_multi_building_with_basement_matches(tmp_path, job, rates):
    from bidpkg.models import Structure
    job.structures.append(Structure("Tower", "x", "p", 5000, 12, 5000, "1988", "o", "", "s"))
    p = tmp_path / "t.xlsx"; write_workbook(job, rates, p); assert recalc(p) == {}
    v = openpyxl.load_workbook(p, data_only=True)
    assert summary_value(v, "Crew & time price") == price_crew_time(job, rates).total
    assert summary_value(v, "$/SF price") == price_per_sf(job, rates).total

def test_open_questions_and_price_to_send(built):
    recalc(built)
    v = openpyxl.load_workbook(built, data_only=True)
    assert summary_value(v, "Open questions") == 1
    assert summary_value(v, "Price to send") is None
