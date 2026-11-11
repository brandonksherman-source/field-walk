import subprocess, sys
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]
FIXTURE_JOB = ROOT / "tests" / "fixtures" / "fixture_job.json"

def run(args):
    r = subprocess.run([sys.executable, "-m", "bidpkg.cli", *map(str, args)], cwd=ROOT, capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    return r.stdout

def test_build_makes_all_outputs(tmp_path):
    run(["init", "--root", tmp_path])
    out = run(["build", FIXTURE_JOB, "--root", tmp_path, "--rates", tmp_path / "Ahrens Rate Card.xlsx"])
    folder = Path(out.strip().splitlines()[-1])
    names = {p.name for p in folder.iterdir()}
    assert {"photos"} | {f"Fixture Bldg - {s}" for s in ("Takeoff.xlsx", "Proposal.docx", "Proposal.pdf", "Bid Intel.md")} <= names
    intel = (folder / "Fixture Bldg - Bid Intel.md").read_text()
    assert "Crew & time price" in intel and "$/SF price" in intel and "No hazmat survey provided" in intel
    assert "Bid due: NOT GIVEN" in intel
    import openpyxl
    assert openpyxl.load_workbook(tmp_path / "Bid Log.xlsx").active.max_row == 2

def test_stop_writes_only_note(tmp_path):
    run(["stop", "could not identify the building", "--root", tmp_path, "--client", "Unknown", "--project", "Photo 11-11"])
    f = next((tmp_path / "1 - Ready for Review").iterdir())
    assert [p.name for p in f.iterdir() if p.is_file()] == ["STOPPED – could not identify the building.md"]

def test_init_never_overwrites(tmp_path):
    run(["init", "--root", tmp_path]); rc = tmp_path / "Ahrens Rate Card.xlsx"; m = rc.stat().st_mtime_ns
    run(["init", "--root", tmp_path]); assert rc.stat().st_mtime_ns == m
    assert {p.name for p in tmp_path.iterdir()} >= {"1 - Ready for Review", "2 - Sent", "3 - Passed", "Past Estimates", "Bid Log.xlsx"}
