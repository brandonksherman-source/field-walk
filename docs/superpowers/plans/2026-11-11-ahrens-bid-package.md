# Ahrens Bid Package Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A Claude skill that takes an address, photo, email/flyer or bid package for a demolition job and produces a review-ready Ahrens bid package (takeoff workbook, proposal letter .docx/.pdf, bid-intel summary) in one job folder, priced two ways with per-line include switches.

**Architecture:** Claude does the judgment work (reading inputs, web research, choosing floors, writing scope wording) by following `SKILL.md`. Deterministic work lives in a small Python package `bidpkg` that the skill calls: map/parcel lookups, quantities, both pricing methods, the workbook, the letter, the job log and past-estimate calibration. Everything is written to a local job folder first, then copied to Drive with the Google Drive connector, so moving off Drive later only changes that last step.

**Tech Stack:** Python 3.13, openpyxl 3.1, python-docx, Pillow (EXIF), LibreOffice headless (`soffice`) for PDF and workbook recalculation, pytest, `urllib` for OSM/Nominatim/St. Louis ArcGIS.

**Spec:** `docs/superpowers/specs/2026-11-11-ahrens-bid-package-design.md`

## Global Constraints

- All outputs are ordinary files (.xlsx, .docx, .pdf, .md). No Google-native Docs or Sheets.
- The skill never sends, emails, submits or uploads a bid anywhere except the `Ahrens Bids/` review folders.
- Every number carries its source. Placeholder rates are shown in blue text on yellow fill and labeled `market placeholder – replace with Ahrens rates/quotes`.
- Workbook color legend: blue = input, black = formula, green = cross-sheet link, yellow fill = verify/placeholder. Font Arial.
- Full scope by default: every scope line `Include? = Y`; scrap credit to Ahrens by default (subtracted).
- A line set to `N` is zeroed in both prices and listed under Exclusions; a line set to `Y` is an Items Included bullet and is never listed under Exclusions.
- Fixed conditions always listed as exclusions: SWPP maintenance and reports after demolition is complete; Pavement repair or damage; Permitting beyond our work; Additional cost for "wet demolition"; All work to be completed during 7-3:30. If work requires labor and/or equipment beyond these hours, an additional charge will be incurred.
- The letter has no signature and no person's name. Closing: "If you have any questions, please contact us at 314-631-7799." / "Best Regards," / blank space / "Ahrens Contracting, Inc." / disclaimer.
- Letterhead text: "Ahrens Contracting, Inc." / "140 Lafayette Ave." / "St. Louis, MO 63104" / "Ph 314-631-7799   Fax 314-621-0977" / "WBE, WBENC, Avetta & ISN Certified". Footer: "Ahrens Contracting Inc. • 140 Lafayette Avenue • St. Louis, MO 63104".
- Opening line: "Ahrens Contracting will provide the labor, material and equipment to complete the following scope of work."
- Lump sum shown as `LUMP SUM: $73,777.00` style (bold, right-aligned, whole dollars, two decimals).
- Drafts carry a "DRAFT – for review" watermark/header.
- HTTP requests send `User-Agent: ahrens-bid-research/0.1` and no personal details.
- Drive layout: `Ahrens Bids/1 - Ready for Review/<YYYY-MM-DD> <Client> - <Project>/`, `Ahrens Bids/2 - Sent`, `Ahrens Bids/3 - Passed`, `Ahrens Bids/Ahrens Rate Card.xlsx`, `Ahrens Bids/Bid Log.xlsx`, `Ahrens Bids/Past Estimates/`.

## Review Focus

1. **Project or client names with `/ : & ' "`** (e.g. "Building 121/123"): folder and file names must still be valid; expect `/` → `-` and other illegal characters dropped. Test in Task 2.
2. **Every switch off, or scrap credit bigger than the rest:** the lump sum must not go negative; expect it floored at $0 with a `price_warning`. Test in Task 3.
3. **No recipient given:** the letter must still render, with a highlighted `[RECIPIENT]` and `Dear [FIRST NAME],`. Test in Task 5.
4. **Address that geocodes to nothing, or a lot with no OSM building:** expect an empty list and a `STOPPED – <reason>.md` note in the job folder, never a made-up footprint. Tests in Tasks 6 and 9.
5. **Rate card with a blank or text rate cell** (someone clears a yellow cell while pasting real numbers): expect a clear error naming the cell, not a silent $0. Test in Task 2.

---

## File Structure

```
skills/ahrens-bid-package/
  SKILL.md                      # the workflow Claude follows (Task 9)
  requirements.txt              # openpyxl, python-docx, Pillow, pytest
  bidpkg/
    __init__.py
    models.py                   # dataclasses shared by all modules (Task 1)
    scope.py                    # scope-line catalog + wording (Task 1)
    ratecard.py                 # create / load the rate card (Task 2)
    paths.py                    # safe names, job folder layout (Task 2)
    pricing.py                  # quantities -> both prices (Task 3)
    workbook.py                 # takeoff .xlsx (Task 4)
    letter.py                   # proposal .docx/.pdf (Task 5)
    research.py                 # geocode, OSM footprints, STL parcels, photo GPS (Task 6)
    joblog.py                   # Bid Log.xlsx (Task 7)
    calibrate.py                # past estimates -> rates (Task 8)
    cli.py                      # `python -m bidpkg.cli build job.json` (Task 9)
  templates/
    letterhead_logo.png         # cropped from Brandon's screenshot until the original arrives
  tests/
    test_scope.py test_ratecard.py test_paths.py test_pricing.py test_workbook.py
    test_letter.py test_research.py test_joblog.py test_calibrate.py test_cli.py
    fixtures/
```

Run all tests from `skills/ahrens-bid-package/`: `python3 -m pytest -q`.

---

### Task 1: Shared models and the scope-line catalog

**Files:**
- Create: `skills/ahrens-bid-package/requirements.txt`, `bidpkg/__init__.py`, `bidpkg/models.py`, `bidpkg/scope.py`
- Test: `tests/test_scope.py`

**Interfaces:**
- Produces:
  - `models.Structure(name: str, address: str, parcel: str, footprint_sf: float, floors: float, basement_sf: float, year_built: str, owner: str, note: str, source: str)`, property `gross_sf = footprint_sf*floors + basement_sf` unless `published_gsf: float|None` is set, then `published_gsf + basement_sf`.
  - `models.Job(project: str, client: str, recipient: str|None, address: str, date: datetime.date, building_type: str, structures: list[Structure], to_remain: list[str], switches: dict[str,bool], quantities_override: dict[str,float], questions: list[tuple[str,str,str]], sources: list[str], included_bullets_extra: list[str])`
  - `models.Line(key: str, label: str, qty: float, unit: str, rate: float, extended: float, included: bool, rate_source: str)`
  - `models.PriceResult(method: str, lines: list[Line], direct: float, markups: dict[str,float], total: float, warning: str|None)`
  - `scope.SCOPE_LINES: dict[str, ScopeLine]`, where `ScopeLine(key, label, included_bullet: str, exclusion_text: str)`, keys exactly: `mobilization, permits, fencing, utilities, tap_destroys, hazmat, demolition, haul, disposal, scrap_credit, backfill, seed_straw, trees, swppp`
  - `scope.FIXED_EXCLUSIONS: list[str]` (the five strings in Global Constraints, in that order)
  - `scope.default_switches() -> dict[str,bool]` (all True)
  - `scope.included_bullets(switches) -> list[str]`, `scope.exclusions(switches) -> list[str]` (switched-off lines' `exclusion_text`, then `FIXED_EXCLUSIONS`)

- [ ] **Step 1:** `pip install -r requirements.txt` (pytest is not preinstalled). Expected: installs cleanly.
- [ ] **Step 2: Write the failing tests**

```python
def test_default_all_on_and_only_fixed_exclusions():
    sw = default_switches()
    assert all(sw.values()) and set(sw) == set(SCOPE_LINES)
    assert exclusions(sw) == FIXED_EXCLUSIONS

def test_switch_off_moves_line_to_exclusions():
    sw = default_switches(); sw["disposal"] = False; sw["scrap_credit"] = False
    ex = exclusions(sw)
    assert "Tipping or disposal fees (paid by owner)" in ex
    assert "All salvage generated (paid to owner)" in ex
    assert SCOPE_LINES["disposal"].included_bullet not in included_bullets(sw)

def test_hazmat_exclusion_wording_matches_121_123_letter():
    assert SCOPE_LINES["hazmat"].exclusion_text.startswith(
        "All environmental work, including, but not limited to: asbestos abatement, Freon evacuation, bulbs/ballasts")

def test_structure_gross_sf_uses_published_when_set():
    s = Structure("E","2700 Pine","10987000010",56749,8,56749,"1998","A G Edwards","", "OSM", published_gsf=405000)
    assert s.gross_sf == 461749
```

- [ ] **Step 3:** Run `python3 -m pytest tests/test_scope.py -q`. Expected: FAIL (import error).
- [ ] **Step 4:** Implement the models and catalog. Exclusion wording comes from the 121/123 letter where it exists ("All electrical or mechanical isolations or re-routes", "All tap destroys (if required)", "Site fencing and/or security", "Seed/straw & site restoration", "Backfill", "Removal of trees and/or stumps", "Hauling and trucks for salvage (paid by owner)" for `haul`). Included bullets read like the letter's ("Removal and loading of all masonry, brick and CMU block", "Haul all debris to landfill, disposal included", …).
- [ ] **Step 5:** Run the tests. Expected: PASS.
- [ ] **Step 6:** Commit `feat(bidpkg): models and scope-line catalog`.

---

### Task 2: Rate card and safe job paths

**Files:**
- Create: `bidpkg/ratecard.py`, `bidpkg/paths.py`
- Test: `tests/test_ratecard.py`, `tests/test_paths.py`

**Interfaces:**
- Consumes: `scope.SCOPE_LINES` keys.
- Produces:
  - `ratecard.create_rate_card(path: Path) -> None`: one sheet `Rates`, columns `Key | Item | Unit | Rate | Source | Placeholder?`, row 1 title, row 2 legend, header on row 4.
  - `ratecard.load_rates(path: Path) -> dict[str, Rate]`, where `Rate(value: float, unit: str, source: str, placeholder: bool)`. Raises `RateCardError("Rates!D<row> (<key>) is blank or not a number")`.
  - `paths.safe_name(s: str) -> str`, `paths.job_folder(root: Path, date, client, project) -> Path` (creates `<root>/1 - Ready for Review/<YYYY-MM-DD> <Client> - <Project>/` and `photos/`).

Rate keys and placeholder values to seed (all `placeholder=True`): the bid-takeoff table (`mobilization 6500 LS, permit_per_bldg 350, fencing_lf 6.00, erosion_ea 2500, hazmat_survey_per_bldg 1200, abatement_allowance_per_bldg 8000, universal_waste_sf 0.35, electric_disconnect 350, gas_cut_cap 2500, sewer_cap 2800, water_kill 2000, tap_destroy 1500, mech_demo_sf 3.50, highrise_premium_sf 4.00, haul_load 375, tons_per_load 20, disposal_ton 35, scrap_credit_ton 180, fill_lcy 18, place_compact_cy 8, seed_straw_sf 0.10, tree_ea 900, swppp_ls 2500`), crew day rates (`operator_day 720, laborer_day 480, foreman_day 800, driver_day 560`), equipment day rates (`excavator_day 1400, excavator_large_day 2600, highreach_day 4500, skidsteer_day 450, loader_day 900`), production (`sf_per_crew_day_low 6000, sf_per_crew_day_mid 3500, sf_per_crew_day_high 1500` for 1–3, 4–8, 9+ floors), $/SF by type (`sfrate_residential 6.00, sfrate_masonry 7.50, sfrate_steel 9.00, sfrate_concrete 11.00, sfrate_highrise 14.00`), debris (`debris_t_per_sf 0.05, steel_t_per_sf 0.0055`), markups (`overhead_pct 0.10, profit_pct 0.10, contingency_pct 0.10, bond_pct 0.015`). Executors may add keys; tests pin only the ones they use.

- [ ] **Step 1: Write the failing tests**

```python
def test_rate_card_round_trip(tmp_path):
    p = tmp_path/"Ahrens Rate Card.xlsx"; create_rate_card(p)
    r = load_rates(p)
    assert r["mech_demo_sf"].value == 3.50 and r["mech_demo_sf"].placeholder
    assert r["bond_pct"].value == 0.015

def test_blank_rate_names_the_cell(tmp_path):
    p = tmp_path/"rc.xlsx"; create_rate_card(p)
    wb = openpyxl.load_workbook(p); ws = wb["Rates"]
    row = next(r for r in range(5, ws.max_row+1) if ws.cell(r,1).value == "haul_load")
    ws.cell(row,4).value = None; wb.save(p)
    with pytest.raises(RateCardError, match=rf"Rates!D{row} \(haul_load\)"):
        load_rates(p)

def test_safe_name_and_folder(tmp_path):
    assert safe_name('Building 121/123: "Demo" & Co.') == "Building 121-123 Demo & Co"
    f = job_folder(tmp_path, date(2026,11,11), "USRA", "WFA Campus C/D/E")
    assert f == tmp_path/"1 - Ready for Review"/"2026-11-11 USRA - WFA Campus C-D-E"
    assert (f/"photos").is_dir()
```

(`&` is legal on Windows and Drive and stays; `: " * ? < > |` and trailing dots are dropped; `/` and `\` become `-`.)

- [ ] **Step 2:** Run the tests. Expected: FAIL.
- [ ] **Step 3:** Implement. Placeholder rate cells: blue font `0000FF`, yellow fill `FFFF00`, Source = `market placeholder – replace with Ahrens rates/quotes`.
- [ ] **Step 4:** Run the tests. Expected: PASS.
- [ ] **Step 5:** Commit `feat(bidpkg): rate card and job paths`.

---

### Task 3: Quantities and both pricing methods

**Files:**
- Create: `bidpkg/pricing.py`
- Test: `tests/test_pricing.py`

**Interfaces:**
- Consumes: `Job`, `Structure`, `Rate`, `SCOPE_LINES`.
- Produces:
  - `quantities(job: Job, rates) -> dict[str, float]` with keys `gross_sf, above_grade_sf, basement_sf, footprint_sf, buildings, debris_tons, steel_tons, haul_loads, backfill_cy, seed_sf, fence_lf, crew_days, equipment_days, max_floors`; values in `job.quantities_override` win. `backfill_cy = basement_sf*12/27*1.25`, `fence_lf = 4*sqrt(footprint_sf)*1.2`, `seed_sf = footprint_sf`, `debris_tons = gross_sf*debris_t_per_sf`, `steel_tons = gross_sf*steel_t_per_sf`, `haul_loads = ceil((debris_tons)/tons_per_load)`, `crew_days = ceil(gross_sf / sf_per_crew_day_<band>)` with the band from `max_floors`.
  - `price_crew_time(job, rates) -> PriceResult` (method `"crew_time"`): demolition = crew_days × (operator+laborer+foreman day) + equipment days × the excavator rate for the band (`highreach_day` when max_floors ≥ 9); other lines priced from their rates; `scrap_credit` is negative.
  - `price_per_sf(job, rates) -> PriceResult` (method `"per_sf"`): demolition = gross_sf × `sfrate_<building_type>`; the other lines the same as crew_time.
  - Both: a line whose switch is off has `included=False, extended=0`. Markups = contingency, overhead, profit applied in sequence on direct; `total` rounded to whole dollars; total < 0 → total 0 and `warning="Price below zero – check scrap credit and switches"`.

- [ ] **Step 1: Write the failing tests** using a fixture job: one 2-floor masonry building, footprint 10,000 SF, no basement, building_type `masonry`, rates from `create_rate_card`.

```python
def test_switch_off_zeroes_line_in_both_methods(job, rates):
    job.switches["disposal"] = False
    for res in (price_crew_time(job, rates), price_per_sf(job, rates)):
        d = next(l for l in res.lines if l.key == "disposal")
        assert d.extended == 0 and not d.included

def test_scrap_credit_is_negative(job, rates):
    s = next(l for l in price_crew_time(job, rates).lines if l.key == "scrap_credit")
    assert s.extended == pytest.approx(-20000*0.0055*180)

def test_per_sf_demolition(job, rates):
    d = next(l for l in price_per_sf(job, rates).lines if l.key == "demolition")
    assert d.extended == pytest.approx(20000*7.50)

def test_total_never_negative(job, rates):
    for k in job.switches: job.switches[k] = k == "scrap_credit"
    r = price_crew_time(job, rates)
    assert r.total == 0 and "below zero" in r.warning

def test_override_wins(job, rates):
    job.quantities_override["crew_days"] = 15
    assert quantities(job, rates)["crew_days"] == 15

def test_highrise_uses_highreach(job, rates):
    job.structures[0].floors = 12
    d = next(l for l in price_crew_time(job, rates).lines if l.key == "demolition")
    assert "highreach" in d.rate_source
```

- [ ] **Step 2:** Run. Expected: FAIL.
- [ ] **Step 3:** Implement.
- [ ] **Step 4:** Run. Expected: PASS.
- [ ] **Step 5:** Commit `feat(bidpkg): quantities and two pricing methods`.

---

### Task 4: Takeoff workbook

**Files:**
- Create: `bidpkg/workbook.py`
- Test: `tests/test_workbook.py`

**Interfaces:**
- Consumes: `Job`, `quantities`, `price_crew_time`, `price_per_sf`, rates.
- Produces: `write_workbook(job, rates, path: Path) -> None` and `recalc(path: Path) -> dict[str, list[str]]` (runs `soffice --headless --convert-to xlsx` into a temp dir, reopens with `data_only=True`, returns error cells per sheet; empty means clean).

Tabs, in order: `Summary, SF Takeoff, Quantities, Crew & Time, $-SF Check, Bid Checklist, Sources`. Rows 1–3 of each tab: title `<tab> – <project> (<address>)`, request line (client, date), legend. `Crew & Time` and `$-SF Check` have columns `Include? | Line | Qty | Unit | Rate | Extended | Rate source`, with `Include?` a Y/N data-validation dropdown and `Extended = IF(A<r>="Y", C<r>*E<r>, 0)`. Summary shows both totals as green links, their difference, `$ per gross SF`, a blank yellow `Price to send` cell, open-question count, and warnings. `SF Takeoff` mirrors the Wells Fargo columns: Structure | Address/Parcel | Footprint | Floors | Above-grade gross | Basement SF | Total gross | Year | Owner | Demo note | Source, plus a "To remain / protect" block.

- [ ] **Step 1: Write the failing tests**

```python
def test_tabs_and_clean_recalc(tmp_path, job, rates):
    p = tmp_path/"t.xlsx"; write_workbook(job, rates, p)
    assert openpyxl.load_workbook(p).sheetnames == ["Summary","SF Takeoff","Quantities","Crew & Time","$-SF Check","Bid Checklist","Sources"]
    assert recalc(p) == {}

def test_workbook_totals_match_python(tmp_path, job, rates):
    p = tmp_path/"t.xlsx"; write_workbook(job, rates, p); recalc(p)
    v = openpyxl.load_workbook(p, data_only=True)
    assert summary_value(v, "Crew & time price") == price_crew_time(job, rates).total
    assert summary_value(v, "$/SF price") == price_per_sf(job, rates).total

def test_flipping_include_cell_changes_total(tmp_path, job, rates):
    p = tmp_path/"t.xlsx"; write_workbook(job, rates, p)
    set_include(p, "Crew & Time", "disposal", "N"); recalc(p)
    job.switches["disposal"] = False
    assert summary_value(openpyxl.load_workbook(p, data_only=True), "Crew & time price") == price_crew_time(job, rates).total
```

(`summary_value` and `set_include` are small helpers in the test file that find a row by its label in column B or the line key in a hidden column H.)

- [ ] **Step 2:** Run. Expected: FAIL.
- [ ] **Step 3:** Implement with formulas throughout (markups and rounding in-sheet as well, so a reviewer's edits flow through). Never write a text cell starting with `=`.
- [ ] **Step 4:** Run. Expected: PASS.
- [ ] **Step 5:** Commit `feat(bidpkg): takeoff workbook with include switches`.

---

### Task 5: Proposal letter (.docx and .pdf)

**Files:**
- Create: `bidpkg/letter.py`, `templates/letterhead_logo.png` (cropped from `/root/.claude/uploads/1d93e1bc-6ac2-5582-9ef5-7980b3c0815c/1b5f828f-image.jpg`, page 2's letterhead logo, straightened, white background)
- Test: `tests/test_letter.py`

**Interfaces:**
- Consumes: `Job`, `scope.included_bullets`, `scope.exclusions`.
- Produces: `render_letter(job: Job, lump_sum: float, out_dir: Path) -> tuple[Path, Path]` that writes `<Project> - Proposal.docx` and `.pdf` (PDF via `soffice --headless --convert-to pdf`).

The document is built in code with python-docx (no external template until the original file arrives; then Task 10 swaps it in): Calibri 11, 1" margins, the letterhead (logo left, five text lines right) in the header of every page, footer line gray and centered, `DRAFT – for review` in red in the header's right corner. Body order exactly as in Global Constraints and the spec: date ("November 11, 2026"), `To: <recipient>`, `RE: <project> Demo`, `Dear <first name>,`, opening line, underlined `Items Included:`, bullets (`included_bullets(job.switches) + job.included_bullets_extra`), right-aligned bold `LUMP SUM: $X`, bold `Exclusions:`, bullets, closing, small-print disclaimer:

> This proposal is contingent upon and shall not be binding or valid unless and until acceptable subcontract terms are agreed to in writing by Ahrens Contracting, Inc. (which acceptance may be withheld for any reason and the submission of this proposal is not an acceptance of any tendered or proposed subcontract terms); including, but not limited to, terms with respect to payment, schedule for the work and disputes.  The party receiving this proposal may not rely upon this proposal unless and until this contingency is satisfied.

- [ ] **Step 1: Write the failing tests**

```python
def test_letter_content(tmp_path, job):
    docx_p, pdf_p = render_letter(job, 73777, tmp_path)
    text = all_text(docx_p)   # paragraphs + header + footer
    assert "LUMP SUM: $73,777.00" in text
    assert "Items Included:" in text and "Exclusions:" in text
    assert "WBE, WBENC, Avetta & ISN Certified" in text
    assert "Roger Kent" not in text and "VP of Operations" not in text
    assert "DRAFT – for review" in text
    assert pdf_p.exists() and pdf_p.stat().st_size > 5000

def test_switched_off_line_is_excluded_not_included(tmp_path, job):
    job.switches["backfill"] = False
    text = all_text(render_letter(job, 1000, tmp_path)[0])
    inc, exc = text.split("Exclusions:")
    assert "Backfill" not in inc and "Backfill" in exc

def test_missing_recipient_placeholder(tmp_path, job):
    job.recipient = None
    text = all_text(render_letter(job, 1000, tmp_path)[0])
    assert "To: [RECIPIENT]" in text and "Dear [FIRST NAME]," in text
```

- [ ] **Step 2:** Run. Expected: FAIL.
- [ ] **Step 3:** Implement. The `[RECIPIENT]` / `[FIRST NAME]` runs get yellow highlight.
- [ ] **Step 4:** Run. Expected: PASS. Then render the fixture letter to PNG (`soffice --convert-to png` or `pdftoppm`) and compare by eye with the two screenshots; fix spacing until the layout matches.
- [ ] **Step 5:** Commit `feat(bidpkg): Ahrens proposal letter`.

---

### Task 6: Research helpers (geocode, OSM footprints, St. Louis parcels, photo GPS)

**Files:**
- Create: `bidpkg/research.py`
- Test: `tests/test_research.py` (offline tests always; live tests marked `@pytest.mark.network`, run with `-m network`)

**Interfaces:**
- Produces:
  - `polygon_area_sf(coords: list[tuple[float,float]]) -> float` ((lon, lat) ring; same equirectangular formula as bid-takeoff's JS).
  - `geocode(address: str) -> list[dict]` (Nominatim; `lat, lon, display_name`; `[]` when nothing).
  - `osm_buildings(lat: float, lon: float, half_box_deg: float = 0.002) -> list[dict]` (`osm_type, osm_id, name, building, levels, footprint_sf`), re-fetching any way/relation that crosses the box via `/full.json`; relations = outer minus inner.
  - `stl_parcels(lat, lon) -> list[dict]` (St. Louis `St_Louis_Parcels/MapServer/0/query` with a point geometry; returns `SITEADDR, OWNERNAME, AsrLandUse1, FirstYearBuilt, LastYearBuilt, NbrOfBldgsCom, NbrOfBldgsRes, SQFT, CityBlock, HANDLE`); `[]` outside the city.
  - `photo_gps(path: Path) -> tuple[float,float] | None` (EXIF GPS via Pillow).
- All HTTP through one `_get_json(url)` with the User-Agent from Global Constraints, a 20 s timeout and one retry; 1 request/second to Nominatim.

- [ ] **Step 1: Write the failing tests**

```python
def test_area_of_100ft_square():
    # 100 ft ≈ 30.48 m; at lat 38.63 build a square of that size in degrees
    assert polygon_area_sf(square_ring(38.63, -90.21, 30.48)) == pytest.approx(10000, rel=0.01)

def test_photo_without_gps(tmp_path):
    p = tmp_path/"x.jpg"; Image.new("RGB",(10,10)).save(p)
    assert photo_gps(p) is None

def test_photo_with_gps(tmp_path):
    p = make_gps_jpeg(tmp_path, 38.6353, -90.2205)   # helper writes GPS IFD with Pillow's Exif API
    assert photo_gps(p) == pytest.approx((38.6353, -90.2205), abs=1e-4)

@pytest.mark.network
def test_wells_fargo_building_e_live():
    lat, lon = float(geocode("2700 Pine St, St. Louis, MO")[0]["lat"]), float(geocode("2700 Pine St, St. Louis, MO")[0]["lon"])
    assert any(p["OWNERNAME"].upper().startswith("A G EDWARDS") for p in stl_parcels(lat, lon))
    assert max(b["footprint_sf"] for b in osm_buildings(lat, lon)) > 40000

@pytest.mark.network
def test_nonsense_address_live():
    assert geocode("zzzz no such place 99999") == []
```

- [ ] **Step 2:** Run `python3 -m pytest tests/test_research.py -q -m "not network"`, then `-m network`. Expected: FAIL.
- [ ] **Step 3:** Implement.
- [ ] **Step 4:** Run both. Expected: PASS.
- [ ] **Step 5:** Commit `feat(bidpkg): map, parcel and photo research helpers`.

---

### Task 7: Bid log

**Files:**
- Create: `bidpkg/joblog.py`
- Test: `tests/test_joblog.py`

**Interfaces:**
- Produces: `append_job(log_path: Path, job: Job, crew: PriceResult, per_sf: PriceResult, folder: Path) -> int` (creates the log if missing, returns the row). Columns: `Date | Client | Project | Address | Gross SF | Crew & time price | $/SF price | Price sent | Won/Lost | Actual cost | Folder`. The last four are left blank for the team (Price sent, Won/Lost and Actual cost are yellow).

- [ ] **Step 1: Write the failing test**

```python
def test_append_creates_then_appends(tmp_path, job, rates):
    log = tmp_path/"Bid Log.xlsx"
    a, b = price_crew_time(job, rates), price_per_sf(job, rates)
    assert append_job(log, job, a, b, tmp_path) == 2
    assert append_job(log, job, a, b, tmp_path) == 3
    ws = openpyxl.load_workbook(log).active
    assert ws["F2"].value == a.total and ws["H2"].value is None
```

- [ ] **Step 2:** Run. Expected: FAIL. **Step 3:** Implement. **Step 4:** Run. Expected: PASS.
- [ ] **Step 5:** Commit `feat(bidpkg): bid log`.

---

### Task 8: Calibration from past estimates

**Files:**
- Create: `bidpkg/calibrate.py`
- Test: `tests/test_calibrate.py`

Claude reads each past estimate (PDF/Word/photo) and measures each building, writing one row per job to `past_jobs.csv`: `job_id, date, client, project, address, building_type, gross_sf, max_floors, lump_sum, crew_weeks, crew_size, scope_included (semicolon list of scope keys), won (Y/N/blank), actual_cost, source_file, located (Y/N)`. Python does the math.

**Interfaces:**
- Produces:
  - `load_past_jobs(csv_path: Path) -> list[dict]`
  - `calibrate(jobs) -> dict[str, dict]` keyed by rate key (`sfrate_<type>`, and `sf_per_crew_day_<band>` where crew_weeks is known): `{"value": median, "n": count, "jobs": [job_id,...], "spread": (min, max)}`. Uses only `located == "Y"` rows; skips any type with n < 3.
  - `write_calibration(result, jobs, path: Path) -> None` (`Calibration.xlsx`: `Jobs` tab and `Derived Rates` tab).
  - `apply_to_rate_card(result, rate_card: Path) -> list[str]` (updates those rates, sets Source to `Ahrens past jobs (n=<n>): <ids>`, clears yellow, returns changed keys).

- [ ] **Step 1: Write the failing tests**

```python
def test_median_sf_rate_by_type(rows):   # 3 masonry jobs: $/SF 6.0, 7.0, 9.0; 2 steel jobs
    r = calibrate(rows)
    assert r["sfrate_masonry"]["value"] == 7.0 and r["sfrate_masonry"]["n"] == 3
    assert "sfrate_steel" not in r

def test_unlocated_rows_ignored(rows):
    rows[0]["located"] = "N"
    assert "sfrate_masonry" not in calibrate(rows)   # 2 located masonry jobs < minimum of 3

def test_apply_updates_source_and_clears_placeholder(tmp_path, rows):
    rc = tmp_path/"rc.xlsx"; create_rate_card(rc)
    changed = apply_to_rate_card(calibrate(rows), rc)
    r = load_rates(rc)["sfrate_masonry"]
    assert "sfrate_masonry" in changed and r.value == 7.0 and not r.placeholder and r.source.startswith("Ahrens past jobs (n=3)")
```

($/SF for a past job = lump_sum ÷ gross_sf, computed only over jobs whose `scope_included` has `demolition`; the plan leaves normalizing for scope mix to a later pass once real data shows whether it's needed.)

- [ ] **Step 2:** Run. Expected: FAIL. **Step 3:** Implement. **Step 4:** Run. Expected: PASS.
- [ ] **Step 5:** Commit `feat(bidpkg): calibrate rates from past estimates`.

---

### Task 9: Command line, SKILL.md and the Drive hand-off

**Files:**
- Create: `bidpkg/cli.py`, `SKILL.md`
- Test: `tests/test_cli.py`

**Interfaces:**
- Consumes: everything above.
- Produces:
  - `python3 -m bidpkg.cli build <job.json> --root <dir> --rates <rate card>`: loads `Job` from JSON (same field names as `models.Job`), makes the job folder, writes the workbook, letter (priced at crew-and-time) and `<Project> - Bid Intel.md`, appends the bid log, prints the folder path. Exit code 0.
  - `python3 -m bidpkg.cli stop <reason> --root <dir> --client C --project P`: makes the folder and writes `STOPPED – <reason>.md` only. Exit code 0.
  - `python3 -m bidpkg.cli init --root <dir>`: creates `Ahrens Rate Card.xlsx`, `Bid Log.xlsx` and the three queue folders plus `Past Estimates/` if missing; never overwrites.
  - `python3 -m bidpkg.cli calibrate <past_jobs.csv> --rates <rate card> --out <Calibration.xlsx>`

- [ ] **Step 1: Write the failing tests**

```python
def test_build_makes_all_outputs(tmp_path):
    run(["init","--root",tmp_path]); out = run(["build", FIXTURE_JOB, "--root", tmp_path, "--rates", tmp_path/"Ahrens Rate Card.xlsx"])
    folder = Path(out.strip())
    names = {p.name for p in folder.iterdir()}
    assert {"photos"} | {f"Fixture Bldg - {s}" for s in ("Takeoff.xlsx","Proposal.docx","Proposal.pdf","Bid Intel.md")} <= names

def test_stop_writes_only_note(tmp_path):
    run(["stop","could not identify the building","--root",tmp_path,"--client","Unknown","--project","Photo 11-11"])
    f = next((tmp_path/"1 - Ready for Review").iterdir())
    assert [p.name for p in f.iterdir() if p.is_file()] == ["STOPPED – could not identify the building.md"]

def test_init_never_overwrites(tmp_path):
    run(["init","--root",tmp_path]); rc = tmp_path/"Ahrens Rate Card.xlsx"; m = rc.stat().st_mtime_ns
    run(["init","--root",tmp_path]); assert rc.stat().st_mtime_ns == m
```

- [ ] **Step 2:** Run. Expected: FAIL. **Step 3:** Implement `cli.py`. **Step 4:** Run. Expected: PASS.
- [ ] **Step 5: Write `SKILL.md`.** Frontmatter `name: ahrens-bid-package`, description: "Use when Brandon gives an address, building photo, forwarded email/flyer or bid package for a demolition job and wants a review-ready Ahrens bid (takeoff, proposal letter, summary)." Body = the spec's nine steps in order, telling Claude to:
  1. run every step without asking; ask only if the building can't be identified (then `cli stop`).
  2. use `bidpkg.research` for geocode/parcels/footprints (no browser needed) and web search for floors, published SF, history, hazmat and permit rules; cite sources;
  3. choose `building_type` and floors; for a stepped building enter each tier as its own `Structure` row
     (e.g. "D – tier 18 fl", with its share of the footprint) like the Wells Fargo Tower D tab; flag source
     disagreement > 15%; if a Field Walk export (`field-walk-*.json`) is supplied, use its photos and checklist
     answers and say what they changed;
  4. write `job.json` and run `cli build`;
  5. download `Ahrens Rate Card.xlsx` from Drive `Ahrens Bids/` before building and upload the whole job folder and updated `Bid Log.xlsx` back with the Google Drive connector (create the folder tree on first run with `cli init`, then upload);
  6. reply with the crew-and-time price, the $/SF price, what drives the difference, gross SF per structure, due date, top 3–5 risks, placeholder rates used, and the Drive folder link.
  Rules section: never send/submit/email a bid; label every number with its source; no signature or names on the letter; never invent a footprint.
  Also a short "Calibrate from past estimates" section: read each file in `Ahrens Bids/Past Estimates/`, measure each building with the same research steps, write `past_jobs.csv`, run `cli calibrate`, upload `Calibration.xlsx` and the updated rate card, and report which rates changed and how many jobs each rests on.
- [ ] **Step 6: Verify the Drive hand-off by hand:** run `cli init` into a scratch folder, upload `Ahrens Rate Card.xlsx` to a test folder in Drive with the connector, download it again, and confirm `load_rates` reads it. Record what worked in SKILL.md (exact connector calls).
- [ ] **Step 7:** Commit `feat: ahrens-bid-package skill entry point`.

---

### Task 10: End-to-end acceptance on real jobs

**Files:**
- Create: `tests/fixtures/wells_fargo_job.json` (from the Rev 1 workbook in Drive, id `1B0KpVj90M11zjYdRmrBl4JCjKrEt6rSXcvYv__3YdjI`)
- Modify: `bidpkg/letter.py` only if the original 121/123 file has arrived (use it for fonts/margins/logo)

- [ ] **Step 1:** Following SKILL.md from the address only ("1 N. Jefferson Ave, St. Louis MO 63103 – Buildings C, D & E, USRA, Ryan Fitzgerald"), produce the Wells Fargo package. Expected: total gross SF within 10% of the Rev 1 workbook's 1,311,052 SF; structures E, D, C and connector identified; F, L and the Olive St garage listed as to remain; Bid Checklist includes bid date, garage in/out, basements, hazmat survey, CRO review.
- [ ] **Step 2:** Produce the Building 121/123 package from whatever Brandon provides (photo or address). Expected: the letter's layout matches the screenshots page for page; the included bullets read like the original's.
- [ ] **Step 3:** Run the full test suite: `python3 -m pytest -q && python3 -m pytest -q -m network`. Expected: all pass.
- [ ] **Step 4:** Upload both packages to `Ahrens Bids/1 - Ready for Review/` and send Brandon the links for the team's review.
- [ ] **Step 5:** Commit `test: Wells Fargo acceptance fixture`.
