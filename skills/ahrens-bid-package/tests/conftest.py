import sys, copy
from datetime import date
from pathlib import Path
import pytest
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from bidpkg.models import Job, Structure
from bidpkg.scope import default_switches

def pytest_configure(config):
    config.addinivalue_line("markers", "network: hits live map/parcel services")

def pytest_collection_modifyitems(config, items):
    if "network" not in (config.getoption("-m") or ""):
        skip = pytest.mark.skip(reason="network test; run with -m network")
        for it in items:
            if "network" in it.keywords:
                it.add_marker(skip)

@pytest.fixture
def job():
    return Job(project="Fixture Bldg", client="Acme Dev", recipient="Bruce Byrd",
               address="100 Test St, St. Louis, MO", date=date(2026, 11, 11), building_type="masonry",
               structures=[Structure("Main", "100 Test St", "123", 10000, 2, 0, "1950", "Acme", "", "OSM")],
               to_remain=["Neighbor garage"], switches=default_switches(), quantities_override={},
               questions=[("Bid date", "Not given", "HIGH")], sources=["OSM way 1"], included_bullets_extra=[])

@pytest.fixture
def rates(tmp_path):
    from bidpkg.ratecard import create_rate_card, load_rates
    p = tmp_path / "rc.xlsx"; create_rate_card(p); return load_rates(p)
