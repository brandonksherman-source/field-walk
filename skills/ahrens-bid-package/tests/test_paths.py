from datetime import date
from bidpkg.paths import safe_name, job_folder

def test_safe_name():
    assert safe_name('Building 121/123: "Demo" & Co.') == "Building 121-123 Demo & Co"
    assert safe_name("a\\b*c?d<e>f|g") == "a-bcdefg"

def test_job_folder(tmp_path):
    f = job_folder(tmp_path, date(2026,11,11), "USRA", "WFA Campus C/D/E")
    assert f == tmp_path/"1 - Ready for Review"/"2026-11-11 USRA - WFA Campus C-D-E"
    assert (f/"photos").is_dir()
    assert job_folder(tmp_path, date(2026,11,11), "USRA", "WFA Campus C/D/E") == f
