# Ahrens Bid Package — Design (Part 1 of the bid automation)

Date: 2026-11-11 · Owner: Brandon Sherman (estimating, Ahrens Contracting) · Status: approved by Brandon 11/11/2026

## Why

Brandon's end goal is bidding with almost no hands-on work: a job comes in (an email, or an agent that finds
buildings slated for demolition), and an estimate plus a ready-to-send proposal lands in front of the team in the
same format every time. The team reviews and sends. People stay in the loop only at review.

The Wells Fargo Advisors campus bid (USRA, 11/9–11/10/2026) is the model for the research and takeoff. The Building
121/123 proposal letter (Oct 12, 2026, signed by Roger Kent) is the model for what the customer receives.

This document covers **Part 1: one job in, one review-ready bid package out.** Part 2, the **Finder** (watching the
estimating inbox and public sources for buildings to be torn down and feeding each one to Part 1), gets its own
design once Part 1 produces packages the team accepts.

## Goals

1. Start from the least possible input: an address, a photo of a building, a forwarded email/flyer, or a bid package.
2. Repeat the Wells Fargo research every time: parcel/owner/year, footprints, floors, gross SF, scope in vs. "to
   remain/protect", hazmat and permit flags, and an open-questions checklist.
3. Price every job two ways, side by side:
   - **Crew and time:** crew × days × day rates + equipment days + haul loads, then markup.
   - **$/SF check:** gross SF × a demo rate by building type.
   The reviewer chooses or adjusts the number. Nothing forces one method while the real Ahrens math is still being
   worked out.
4. Produce the **Ahrens proposal letter** in the house format, every time.
5. Drop everything into one job folder in a "Ready for Review" queue, and log the job so estimates can be compared
   with what was sent and, later, with actual cost.
6. Run with no person at the computer (so Part 2 can call it on a schedule).

## Non-goals (Part 1)

- Finding jobs on its own (Part 2).
- Sending anything. It never emails, submits or uploads a bid. It prepares only.
- Final, "correct" Ahrens pricing. The rate card starts with placeholders and the team fills in real numbers.
- A custom app or server (rejected approach B; maybe later).

## Anti-goals (would be a failure even if it "works")

- A letter that looks different from job to job, or doesn't look like the 121/123 letter.
- A price and a letter that disagree: a scope line switched off in the workbook but still in the lump sum, or
  switched on but still listed under Exclusions.
- Presenting a placeholder rate or an estimated floor count as measured or quoted.
- Locking the files into Google Drive in a way that's hard to move off later.

## Constraints

- Brandon wants to move off Google Drive eventually. The office runs Windows, Outlook, Teams, Excel and Acrobat,
  which suggests Microsoft 365 later. So: **all outputs are ordinary files (.xlsx, .docx, .pdf), not Google-native
  Docs/Sheets.** Drive is only the folder they sit in today. Moving later means changing the save step, not the
  skill.
- Real Ahrens rates live in an Excel sheet or Sage. Brandon will find them; until then, placeholders.
- From a cloud session the OpenStreetMap API, Nominatim and the City of St. Louis parcel service all answer directly
  (checked 11/11/2026: HTTP 200), so the research doesn't need Claude in Chrome. That's what makes unattended runs
  possible. The existing `bid-takeoff` skill assumes Chrome; the new skill does not.

## Approaches considered

- **A. Claude skill + Drive folder (chosen).** Built on the existing `bid-takeoff` skill and the connected Gmail,
  Drive and Sheets. Fastest to working; easy to change while the pricing math settles.
- **B. Custom program on a server calling Claude.** More control, but every pricing change becomes a code change
  while the math is still moving, and someone has to maintain it. Revisit once the process is stable.
- **C. Bid Desk page as the review hub.** Better review experience, but more to build up front. Planned as a later
  layer on top of A.

Brandon chose A, with the goal of leaving Drive later.

## What gets built

### 1. The `ahrens-bid-package` skill (successor to `bid-takeoff`)

**Inputs (any one is enough):** street address; photo(s) of the building; a forwarded email, flyer or listing; a
full bid package (ITB, plans, specs). Field Walk JSON exports are accepted when present (site photos, checklist
answers), as Rev 1 of Wells Fargo was revised from site photos.

**Steps (from Wells Fargo):**
1. **Intake:** pull project name, client/contact, dates, submittal rules and scope out of whatever was given. From a
   photo alone: identify the building (EXIF GPS if present, otherwise visible signage/landmarks plus a web search)
   and confirm the address before continuing. If the building can't be pinned down, stop and say so.
2. **Identify structures:** geocode; query the city/county parcel service for owner, land use, year built and
   building count; list every structure in scope and every structure to remain/protect (neighbors, attached
   buildings, garages inside the outline but not named).
3. **Measure:** OSM footprints (±10%), floors from published sources (SKYDB/Emporis, listings, 10-Ks, news),
   published gross SF when it exists. Basement assumed = footprint when unknown (yellow). For stepped/irregular
   buildings, a tiers table like Wells Fargo's Tower D tab. Cross-check two sources and flag >15% disagreement.
4. **Research:** owner/sale history, former uses, hazmat red flags (pre-1980, gas stations, labs, kitchens), local
   permit rules (City of St. Louis certified demo contractor, notarized owner letter, Cultural Resources Office /
   Preservation Board). Cite every source.
5. **Quantities:** debris tons, steel scrap tons, haul loads, crew-days and equipment-days, each formula with its
   input in a blue cell and its assumption written beside it.
6. **Price both ways** (see Rate card), apply markups, show the two totals and the difference.
7. **Scope bullets and exclusions:** generated from the scope switches (see Full-scope pricing). Every line that
   is on becomes an "Items Included" bullet; every line that is off moves to "Exclusions". The fixed conditions
   (below) are always listed as exclusions.
8. **Open questions checklist** for the client (bid date, scope boundaries, basements, hazmat survey, utilities,
   occupancy, permits).
9. **Write outputs** and **log the job** (below).

### Full-scope pricing (Brandon, 11/11)

Every job is bid as if Ahrens does **all** of it, and the lump sum includes all of it:
demolition and loading, trucking/haul, tipping/disposal, utility disconnects and cut-and-caps (electric, gas,
water, sewer), tap destroys, hazmat (survey, abatement, universal waste such as bulbs/ballasts/Freon), fencing and
protection, backfill, seed/straw and site restoration, tree/stump removal where trees are found, permits, SWPPP.

**Scrap credit goes to Ahrens by default** and is subtracted from the price. It's a per-job switch, because on some
jobs the salvage goes to the owner.

Each scope line in the workbook has an **Include? (Y/N)** switch, default Y. Setting a line to N zeroes it in both
prices and moves it to the letter's Exclusions, so the reviewer changes one cell when a customer wants something
handled differently.

### 2. Outputs, per job

One folder: `Ahrens Bids/1 - Ready for Review/<YYYY-MM-DD> <Client> - <Project>/`

- `<Project> - Takeoff.xlsx`: the internal workbook, same tabs and look every time:
  `Summary` (both prices, the chosen price cell left for the reviewer, top risks), `SF Takeoff`, `Quantities`,
  `Crew & Time`, `$-SF Check`, `Bid Checklist`, `Sources`. Same color legend as Wells Fargo: blue input, black
  formula, green cross-sheet link, yellow verify/placeholder.
- `<Project> - Proposal.docx` and `.pdf`: the Ahrens letter (format below), priced at the crew-and-time number
  unless the skill flags that it doesn't fit, with a "DRAFT – for review" watermark the reviewer removes.
- `<Project> - Bid Intel.md`: a one-page summary: price, what drives it, gross SF per structure, due date, the
  3–5 biggest risks, which placeholder rates were used, sources.
- `photos/`: any photos used.

### 3. The proposal letter template

A fixed `.docx` template matching the Building 121/123 letter:

- Letterhead block on each page: logo, Ahrens Contracting, Inc., 140 Lafayette Ave., St. Louis, MO 63104,
  Ph 314-631-7799, Fax 314-621-0977, "WBE, WBENC, Avetta & ISN Certified". Footer: "Ahrens Contracting Inc. •
  140 Lafayette Avenue • St. Louis, MO 63104".
- Date · "To: <name>" · "RE: <project> Demo" · "Dear <first name>,"
- Fixed line: "Ahrens Contracting will provide the labor, material and equipment to complete the following scope
  of work."
- **Items Included:** (underlined heading), bullets (variable)
- **LUMP SUM: $X** (bold, right-aligned)
- **Exclusions:** standard list (below), adjusted per job
- "If you have any questions, please contact us at 314-631-7799." · "Best Regards," · blank space for
  the sender to sign and add their name at review (no signature and no name on drafts, per Brandon) · "Ahrens
  Contracting, Inc." · the subcontract-terms disclaimer paragraph in small print.

Exclusions are built per job:
- **Switched-off scope lines**, worded the way the 121/123 letter words them (e.g. "Tipping or disposal fees (paid
  by owner)", "All salvage generated (paid to owner)", "All environmental work, including, but not limited to:
  asbestos abatement, Freon evacuation, bulbs/ballasts, …").
- **Fixed conditions, always listed:** SWPP maintenance and reports after demolition is complete; pavement repair or
  damage; permitting beyond our work; additional cost for "wet demolition"; all work to be completed during 7–3:30,
  with an additional charge for labor and/or equipment beyond those hours. (Claude's call on which items are
  conditions rather than scope; Brandon can move any of them.)

Variable fields: date, recipient name, first name, RE line, included bullets, lump sum, exclusion edits.

### 4. The rate card

`Ahrens Bids/Ahrens Rate Card.xlsx`: one sheet the skill reads every run. Sections: crew day rates (operator,
laborer, foreman, driver), equipment day rates (excavator sizes, high-reach, skid steer, loader, trucks),
haul ($/load, tons/load), disposal ($/ton by material), scrap credit ($/ton by metal), utility disconnects and
cut-and-caps (each), tap destroys, hazmat (survey per building, abatement and universal-waste allowances),
fencing, backfill and compaction, seed/straw ($/SF or acre), tree/stump removal, $/SF demo rates by building type,
mobilization, permits, markups (overhead, profit, contingency, bond, prevailing-wage uplift).

Seeded with the `bid-takeoff` placeholder rates, in yellow and labeled placeholder. When Brandon finds the Excel
or Sage numbers (a Sage report exported to Excel works), they replace the yellow cells; nothing in the skill changes.

### 4a. Learning from past estimates (added 11/11, before the rate card is filled)

Brandon uploads past Ahrens estimates to `Ahrens Bids/Past Estimates/` (letters, worksheets, Sage printouts, photos;
as-is). For each job: read the lump sum, scope and exclusions; find and measure the building (gross SF, floors,
construction type); note won/lost and actual cost where given. From the set, derive Ahrens' own rates ($/SF by
building type, crew-weeks vs. size, implied markup) and replace the matching placeholders in the rate card, citing
the jobs each number came from. Also collect Ahrens' wording for scope bullets and exclusions by job type, and take
the exact letter template from any Word file in the set. Jobs that can't be located get flagged, not guessed. The
analysis is saved as `Ahrens Bids/Past Estimates/Calibration.xlsx` so it can be rerun as more jobs are added.

### 5. The job log

`Ahrens Bids/Bid Log.xlsx`: one row per job: date, client, project, address, gross SF, crew-and-time price,
$/SF price, price sent (the team fills it in), won/lost, actual cost (later). This is how the real Ahrens math gets
discovered: after 15–20 jobs, comparing estimated with sent shows which method and which rates match the
estimators.

## Decisions made by Claude (Brandon can override)

- Recipient comes from the request, otherwise left as a highlighted blank. The company line under the blank
  signature space stays (it isn't a person's name).
- The letter is priced at the crew-and-time number by default; the $/SF number is a check.
- Queue = folders: `1 - Ready for Review`, `2 - Sent`, `3 - Passed`. The team moves the folder when done.
- The skill is written as files in this repo (`skills/ahrens-bid-package/`) so it's versioned; Brandon uploads it
  to his Claude account skills to use it. It replaces `bid-takeoff` once it's working.
- The Ahrens logo is needed as an image file for the letterhead (taken from an existing proposal PDF or supplied).
- Until the original 121/123 file arrives, the template is rebuilt from Brandon's screenshots of it; the real file
  replaces it when supplied.

- **Where it runs:** a Claude Code cloud session (this environment). It has Python for the .xlsx/.docx/.pdf work,
  direct access to OSM and the city parcel service, and the Gmail and Drive connectors. On-demand runs start from a
  message ("bid 2700 Pine St" plus photos). Part 2 adds a scheduled task that runs the same skill. A run that stops
  (for example, it can't identify the building) still writes its folder with a `STOPPED – <reason>.md` note so the
  team sees it in the queue.
- **Drive location:** `Ahrens Bids/` in Brandon's My Drive, created by the build, shared with the reviewers Brandon
  names. Rate card and Bid Log are created by the build.
- **Reference files:** the Wells Fargo workbooks are in Drive (`WFA_Campus_Bldgs_CDE_Demo_Takeoff`, id
  1q7WzWeXkuziDQPjU2Uh1G-mkW_Ixd6pT_fhvO9ENuxM; `USRA_-_WFA_Campus_CDE_-_Demo_Takeoff_Rev1_-_2026-11-10`, id
  1B0KpVj90M11zjYdRmrBl4JCjKrEt6rSXcvYv__3YdjI). Tab layout and the Tower D tiers table come from these.
- **Test jobs:** Wells Fargo C/D/E and Building 121/123; a third chosen by Brandon when testing starts.
- Crew and equipment day rates start as labeled yellow placeholders (bid-takeoff has none).

## Left to the builder

- Exact crew-days formula (production rates per SF/ton per crew) as a first guess, all inputs in the rate card.
- How to fill the .docx template (python-docx or similar) and render the PDF.
- How the files reach Drive from an unattended run (Drive connector upload vs. a synced folder).
- Wording of scope bullets per building type.

## How we'll know it works

Run it on three past jobs: Wells Fargo C/D/E (address only), Building 121/123 (a photo or address), and one more
Brandon picks. The team should be able to send each letter after changing only the price and small wording, and
the letter should be indistinguishable in format from the 121/123 letter.
