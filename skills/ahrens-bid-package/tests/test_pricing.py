import math, pytest
from bidpkg.pricing import quantities, price_crew_time, price_per_sf

def lines(res, key):
    return [l for l in res.lines if l.key == key]

def test_quantities_basic(job, rates):
    q = quantities(job, rates)
    assert q["gross_sf"] == 20000 and q["footprint_sf"] == 10000 and q["buildings"] == 1
    assert q["debris_tons"] == pytest.approx(1000) and q["steel_tons"] == pytest.approx(110)
    assert q["haul_loads"] == 50 and q["crew_days"] == 4 and q["backfill_cy"] == 0
    assert q["fence_lf"] == pytest.approx(4*100*1.2)

def test_switch_off_zeroes_line_in_both_methods(job, rates):
    job.switches["disposal"] = False
    for res in (price_crew_time(job, rates), price_per_sf(job, rates)):
        for d in lines(res, "disposal"):
            assert d.extended == 0 and not d.included

def test_switch_off_lowers_total(job, rates):
    before = price_crew_time(job, rates).total
    job.switches["hazmat"] = False
    assert price_crew_time(job, rates).total < before

def test_scrap_credit_is_negative(job, rates):
    (s,) = lines(price_crew_time(job, rates), "scrap_credit")
    assert s.extended == pytest.approx(-20000*0.0055*180)

def test_per_sf_demolition(job, rates):
    (d,) = lines(price_per_sf(job, rates), "demolition")
    assert d.extended == pytest.approx(20000*3.50)

def test_crew_time_demolition(job, rates):
    ext = sum(l.extended for l in lines(price_crew_time(job, rates), "demolition"))
    assert ext == pytest.approx(4*(720+480+800) + 4*1400)

def test_markups_and_rounding(job, rates):
    r = price_crew_time(job, rates)
    assert r.total == round(r.direct*1.1*1.1*1.1)
    assert sum(l.extended for l in r.lines) == pytest.approx(r.direct)

def test_total_never_negative(job, rates):
    for k in job.switches: job.switches[k] = k == "scrap_credit"
    r = price_crew_time(job, rates)
    assert r.total == 0 and "below zero" in r.warning

def test_override_wins(job, rates):
    job.quantities_override["crew_days"] = 15
    assert quantities(job, rates)["crew_days"] == 15

def test_highrise_uses_highreach(job, rates):
    job.structures[0].floors = 12
    assert any("highreach" in l.rate_source for l in lines(price_crew_time(job, rates), "demolition"))

def test_basement_drives_backfill(job, rates):
    job.structures[0].basement_sf = 10000
    q = quantities(job, rates)
    assert q["backfill_cy"] == pytest.approx(10000*12/27*1.25)
    (b,) = lines(price_crew_time(job, rates), "backfill")
    assert b.extended == pytest.approx(q["backfill_cy"]*(18+8))
