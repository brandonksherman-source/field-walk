from bidpkg.scope import SCOPE_LINES, FIXED_EXCLUSIONS, default_switches, included_bullets, exclusions
from bidpkg.models import Structure, Job

KEYS = {"mobilization","permits","fencing","utilities","tap_destroys","hazmat","demolition","haul","disposal",
        "scrap_credit","backfill","seed_straw","trees","swppp"}

def test_default_all_on_and_only_fixed_exclusions():
    sw = default_switches()
    assert all(sw.values()) and set(sw) == set(SCOPE_LINES) == KEYS
    assert exclusions(sw) == FIXED_EXCLUSIONS
    assert len(FIXED_EXCLUSIONS) == 5

def test_switch_off_moves_line_to_exclusions():
    sw = default_switches(); sw["disposal"] = False; sw["scrap_credit"] = False
    ex = exclusions(sw)
    assert "Tipping or disposal fees (paid by owner)" in ex
    assert "All salvage generated (paid to owner)" in ex
    assert SCOPE_LINES["disposal"].included_bullet not in included_bullets(sw)
    assert ex[-5:] == FIXED_EXCLUSIONS

def test_hazmat_exclusion_wording_matches_121_123_letter():
    assert SCOPE_LINES["hazmat"].exclusion_text.startswith(
        "All environmental work, including, but not limited to: asbestos abatement, Freon evacuation, bulbs/ballasts")

def test_structure_gross_sf_uses_published_when_set():
    s = Structure("E","2700 Pine","10987000010",56749,8,56749,"1998","A G Edwards","","OSM", published_gsf=405000)
    assert s.gross_sf == 461749
    assert Structure("C","x","p",28888,7,28888,"","","","").gross_sf == 231104

def test_job_round_trips_through_dict(job):
    assert Job.from_dict(job.to_dict()) == job
