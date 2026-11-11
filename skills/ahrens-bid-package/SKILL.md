---
name: ahrens-bid-package
description: "Use when Brandon gives an address, building photo, forwarded email/flyer or bid package for a demolition job and wants a review-ready Ahrens bid (takeoff workbook, proposal letter, summary). Also use to calibrate the Ahrens rate card from past estimates."
---

# Ahrens Bid Package

Brandon is an estimator at Ahrens Contracting (demolition, site work, hauling; St. Louis). One job in → one job
folder out in Drive `Ahrens Bids/1 - Ready for Review/`, the same every time:

- `<Project> - Takeoff.xlsx`: Summary, SF Takeoff, Quantities, Crew & Time, $-SF Check, Bid Checklist, Sources
- `<Project> - Proposal.docx` / `.pdf`: the Ahrens letter, DRAFT, no signature or name
- `<Project> - Bid Intel.md`: one-page summary
- `photos/`, `job.json`

The team reviews, picks the price, signs and sends. **You never send, submit or email a bid.**

Run every step without asking. Ask a question only when you can't tell which building it is.
Set up a task list with the steps below so Brandon can see progress.

## Setup (each run)

Code lives next to this file. Work in a scratch folder `$W`:

```bash
cd <this skill's folder>
pip install -q -r requirements.txt   # openpyxl, python-docx, Pillow, pytest; LibreOffice (soffice) must exist
python3 -m bidpkg.cli init --root "$W"   # makes the queue folders, a blank Bid Log and a seed rate card
```

Then replace the seed files with the live ones from Drive folder `Ahrens Bids` (id `1RIvEGZ9HRwJKu62BBZOJjzqUD0irkW2t`).
Their file ids change whenever they're replaced, so always find them by title:
`Google_Drive.search_files("parentId = '1RIvEGZ9HRwJKu62BBZOJjzqUD0irkW2t' and title = 'Ahrens Rate Card.xlsx'")`
(same for `Bid Log.xlsx`). Note each file's `id` and `modifiedTime`. Download with
`Google_Drive.download_file_content(fileId)` and decode the returned base64 into `$W/<name>` with Python.

## 1. Intake

Read everything supplied: email body, flyer, listing, ITB, plans, specs, photos, Field Walk exports
(`field-walk-*.json`: walk notes, checklist answers, photos with location/heading). Pull out: project name, client
company, contact person (recipient), dates (bid due, walkthrough, questions deadline), submittal method, scope, and
anything marked to remain/protect.

From a photo alone: `research.photo_gps(path)` for the location; otherwise use signage, landmarks and web search.
Confirm the address (geocode it and check the parcel). If the building can't be identified:
`python3 -m bidpkg.cli stop "could not identify the building" --root "$W" --client "<client or Unknown>" --project "<best name>"`,
upload the folder, and tell Brandon what you tried.

## 2. Identify every structure

```python
from bidpkg.research import geocode, stl_parcels, osm_buildings
g = geocode("2700 Pine St, St. Louis, MO")          # [{lat, lon, display_name}] – [] if nothing
stl_parcels(g[0]["lat"], g[0]["lon"])              # City of St. Louis only: owner, year built, land use, parcel
osm_buildings(g[0]["lat"], g[0]["lon"])            # footprints in SF, largest first; widen half_box_deg for campuses
```

List every structure in scope and every structure to remain/protect (attached buildings, garages inside the outline
but not named, neighbors). OSM names often differ from plan names: match by position and parcel. Outside the city
use the county GIS / web.

## 3. Measure

- Footprint: OSM (±10%). Floors: published sources (SKYDB, Emporis, listings, 10-Ks, news), photo window rows.
- Use published gross SF when it exists (`published_gsf`) and say where it came from.
- Basement: assume = footprint when unknown (it shows yellow in the workbook).
- Stepped/irregular building: enter each tier as its own structure ("D – tier 18 fl", footprint share × floors)
  like the Wells Fargo Tower D tab.
- Cross-check two sources; flag > 15% disagreement in `risks`.
- Never invent a footprint. No footprint → say so, use a clearly labeled allowance, or stop.

## 4. Research

Owner and sale history, former uses, hazmat red flags (pre-1980, gas stations, labs, kitchens, boiler houses),
permit rules (City of St. Louis: certified demo contractor, notarized owner letter, Cultural Resources Office /
Preservation Board review). Cite every source.

## 5. Write `job.json` and build

Full scope by default: leave `switches` empty (everything on, scrap credit to Ahrens). Turn a line off only when
the request says the customer handles it (e.g. owner keeps salvage → `"scrap_credit": false`; use true/false). Keys:
`mobilization, permits, fencing, utilities, tap_destroys, hazmat, demolition, haul, disposal, scrap_credit,
backfill, seed_straw, trees, swppp`.

```json
{"project": "WFA Campus Bldgs C, D & E", "client": "USRA LLC", "recipient": "Ryan Fitzgerald",
 "address": "1 N. Jefferson Ave, St. Louis, MO 63103", "date": "2026-11-11", "building_type": "concrete",
 "structures": [{"name": "Building E", "address": "2700 Pine St", "parcel": "10987000010", "footprint_sf": 56749,
   "floors": 8, "basement_sf": 56749, "year_built": "1998", "owner": "A G Edwards & Sons Inc",
   "note": "Linked to Bldg F – cut & close at F face", "source": "OSM rel 19553658; SKYDB; SEC 10-K",
   "published_gsf": 405000}],
 "to_remain": ["Building F & L – protect, separation wall"],
 "switches": {}, "quantities_override": {"trees": 6},
 "questions": [["Bid due date / time", "Not given in 11/9 email", "HIGH"]],
 "sources": ["https://..."], "included_bullets_extra": ["Cut and close skybridge at Olive St garage"],
 "due_date": null, "risks": ["No hazmat survey", "Garage in/out unclear"]}
```

`building_type` is one of `residential, masonry, steel, concrete, highrise` (the $/SF rates in the rate card).
`quantities_override` can set any quantity key (e.g. `trees`, `crew_days`, `basement_sf`).

```bash
python3 -m bidpkg.cli build job.json --root "$W" --rates "$W/Ahrens Rate Card.xlsx"   # prints the job folder
```

## 6. Upload to Drive

Find `1 - Ready for Review` (id `1RLgjQp3D2pUBcbBogKG_gfAbCLmK6Bhb`), create the job folder in it with
`Google_Drive.create_file(title=<folder name>, contentMimeType="application/vnd.google-apps.folder", parentId=…)`,
then upload each file **without conversion**:

```
Google_Drive.create_file(title="<file name>", parentId=<job folder id>,
  contentMimeType=<xlsx: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet |
                   docx: application/vnd.openxmlformats-officedocument.wordprocessingml.document |
                   pdf: application/pdf | md: text/markdown | jpg: image/jpeg>,
  disableConversionToGoogleType=true,
  base64Content="sha256:<sha256sum of the file>:<absolute path>")
```

Replace `Bid Log.xlsx` in `Ahrens Bids` with the updated copy. The connector can't overwrite file contents, so:
look the old file up again; **if its `modifiedTime` changed since you downloaded it, someone edited it: do not
replace it.** Upload yours as `Bid Log (needs merge).xlsx` and tell Brandon. Otherwise upload the new copy to
`Ahrens Bids` and then `Google_Drive.trash_file` the old one (it stays in Trash for 30 days). Same rule for the rate
card after calibration.

## 7. Reply

Briefly:
- **Crew & time price** (on the letter) and **$/SF price**, and what drives the difference
- Gross SF per structure
- Bid due date/time and how long until it's due
- The 3–5 biggest risks or gaps
- How many rates are still placeholders
- The Drive folder link, and sources at the end

## After the team reviews (rebuild the letter)

Reviewers change the workbook, not the letter: they flip `Include?` cells on Crew & Time, edit floors or areas, or
enter **Price to send** on the Summary tab. When asked to update a job ("rebuild the Wells Fargo letter"): download
the job folder's `Takeoff.xlsx` and `job.json` into a local folder, run

```bash
python3 -m bidpkg.cli rebuild "<local job folder>" --root "$W" --rates "$W/Ahrens Rate Card.xlsx"
```

then upload the new `Proposal.docx`/`.pdf`, `job.json` and `Bid Intel.md` (upload new, trash old, same
`modifiedTime` rule). The letter takes Price to send if it's filled in, otherwise the workbook's recalculated crew &
time price, and its Included/Exclusions follow the workbook's switches. Rebuild never adds a Bid Log row.

## Calibrate from past estimates

When asked (or when `Ahrens Bids/Past Estimates` has new files, folder id `1HbdL3hLf4CU_YHc3UIk-800u3fHWZ6Cz`):
1. Read every file. For each job find: date, client, project, address, lump sum, crew size/weeks if stated, which
   scope keys were included, won/lost and actual cost if known.
2. Locate and measure each building with steps 2–3. Set `building_type`, `gross_sf`, `max_floors`.
   Couldn't locate it → `located = N`.
3. Write `past_jobs.csv` with columns `job_id, date, client, project, address, building_type, gross_sf, max_floors,
   lump_sum, crew_weeks, crew_size, scope_included, won, actual_cost, source_file, located` (`scope_included` is a
   semicolon list of scope keys).
4. `python3 -m bidpkg.cli calibrate past_jobs.csv --rates "$W/Ahrens Rate Card.xlsx" --out "$W/Calibration.xlsx"`
5. Upload `Calibration.xlsx` to `Past Estimates` and replace the rate card in Drive.
6. Report which rates changed, from what to what, and how many jobs each rests on. Also collect the scope bullets
   and exclusion wording Ahrens uses by job type and tell Brandon about any that differ from `bidpkg/scope.py`.
   If any past estimate is a Word file in the house format, point it out so the letter template can match it exactly.

## Rules

- Never send, submit, or email a bid. Prepare only.
- Label every number with its source. Never present a placeholder rate or an estimated floor count as measured.
- No signature and no person's name on the letter.
- Never invent a footprint.
- Outputs stay ordinary files (.xlsx/.docx/.pdf/.md). Never convert them to Google Docs or Sheets.
