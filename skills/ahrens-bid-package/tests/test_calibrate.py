import csv, openpyxl, pytest
from bidpkg.calibrate import load_past_jobs, calibrate, write_calibration, apply_to_rate_card, CSV_COLUMNS
from bidpkg.ratecard import create_rate_card, load_rates

def row(i, typ, sf, price, floors=2, weeks="", located="Y", scope="demolition;haul"):
    return {"job_id": f"J{i}", "date": "2026-01-01", "client": "c", "project": f"p{i}", "address": "a",
            "building_type": typ, "gross_sf": str(sf), "max_floors": str(floors), "lump_sum": str(price),
            "crew_weeks": str(weeks), "crew_size": "2", "scope_included": scope, "won": "", "actual_cost": "",
            "source_file": f"f{i}.pdf", "located": located}

@pytest.fixture
def rows():
    return [row(1, "masonry", 10000, 60000, weeks=1), row(2, "masonry", 10000, 70000, weeks=1),
            row(3, "masonry", 10000, 90000, weeks=2), row(4, "steel", 20000, 180000), row(5, "steel", 5000, 50000)]

def test_median_sf_rate_by_type(rows):
    r = calibrate(rows)
    assert r["sfrate_masonry"]["value"] == 7.0 and r["sfrate_masonry"]["n"] == 3
    assert r["sfrate_masonry"]["jobs"] == ["J1", "J2", "J3"] and r["sfrate_masonry"]["spread"] == (6.0, 9.0)
    assert "sfrate_steel" not in r

def test_crew_production_by_band(rows):
    r = calibrate(rows)
    assert r["sf_per_crew_day_low"]["value"] == 2000 and r["sf_per_crew_day_low"]["n"] == 3

def test_unlocated_rows_ignored(rows):
    rows[0]["located"] = "N"
    assert "sfrate_masonry" not in calibrate(rows)   # 2 located masonry jobs < minimum of 3

def test_rows_without_demolition_ignored_for_sf_rate(rows):
    rows[0]["scope_included"] = "haul"
    assert "sfrate_masonry" not in calibrate(rows)

def test_csv_round_trip_and_workbook(tmp_path, rows):
    p = tmp_path / "past_jobs.csv"
    with open(p, "w", newline="") as f:
        w = csv.DictWriter(f, CSV_COLUMNS); w.writeheader(); w.writerows(rows)
    jobs = load_past_jobs(p)
    out = tmp_path / "Calibration.xlsx"; write_calibration(calibrate(jobs), jobs, out)
    assert openpyxl.load_workbook(out).sheetnames == ["Jobs", "Derived Rates"]

def test_apply_updates_source_and_clears_placeholder(tmp_path, rows):
    rc = tmp_path / "rc.xlsx"; create_rate_card(rc)
    changed = apply_to_rate_card(calibrate(rows), rc)
    r = load_rates(rc)["sfrate_masonry"]
    assert "sfrate_masonry" in changed and r.value == 7.0 and not r.placeholder
    assert r.source.startswith("Ahrens past jobs (n=3)")
    assert load_rates(rc)["sfrate_steel"].placeholder
