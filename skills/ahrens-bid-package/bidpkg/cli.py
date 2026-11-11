"""Command line the skill calls.

  python3 -m bidpkg.cli init --root DIR
  python3 -m bidpkg.cli build JOB.json --root DIR --rates "DIR/Ahrens Rate Card.xlsx"
  python3 -m bidpkg.cli stop "REASON" --root DIR --client C --project P
  python3 -m bidpkg.cli calibrate past_jobs.csv --rates RATECARD --out Calibration.xlsx
"""
from __future__ import annotations
import argparse
import json
import sys
from datetime import date
from pathlib import Path

from .models import Job
from .paths import QUEUE, job_folder, safe_name
from .ratecard import create_rate_card, load_rates
from .pricing import price_crew_time, price_per_sf, quantities
from .workbook import write_workbook, recalc
from .letter import render_letter
from .joblog import append_job, create_log
from . import calibrate as cal


def cmd_init(root: Path) -> None:
    root.mkdir(parents=True, exist_ok=True)
    for q in (*QUEUE, "Past Estimates"):
        (root / q).mkdir(exist_ok=True)
    if not (root / "Ahrens Rate Card.xlsx").exists():
        create_rate_card(root / "Ahrens Rate Card.xlsx")
    if not (root / "Bid Log.xlsx").exists():
        create_log(root / "Bid Log.xlsx")
    print(root)


def bid_intel(job: Job, rates, crew, persf) -> str:
    q = quantities(job, rates)
    diff = crew.total - persf.total
    lines = [
        f"# {job.project} – Bid Intel (DRAFT – for review)", "",
        f"**Client:** {job.client} · **To:** {job.recipient or '[RECIPIENT]'} · **Address:** {job.address}",
        f"**Prepared:** {job.date:%m/%d/%Y} · **Bid due: {job.due_date or 'NOT GIVEN – confirm'}**", "",
        "## Price", "",
        f"- **Crew & time price: ${crew.total:,.0f}** (letter uses this)",
        f"- $/SF price: ${persf.total:,.0f} ({job.building_type})",
        f"- Difference: ${diff:,.0f} ({(diff / persf.total * 100) if persf.total else 0:+.0f}%)",
        f"- Total gross SF: {q['gross_sf']:,.0f} · ${crew.total / q['gross_sf'] if q['gross_sf'] else 0:,.2f}/SF",
    ]
    for w in (crew.warning, persf.warning):
        if w:
            lines.append(f"- ⚠️ {w}")
    off = [k for k, v in job.switches.items() if not v]
    lines.append(f"- Switched off (excluded): {', '.join(off) if off else 'none – full scope'}")
    lines += ["", "## Structures", "", "| Structure | Footprint SF | Floors | Gross SF | Year | Source |",
              "|---|---:|---:|---:|---|---|"]
    for s in job.structures:
        lines.append(f"| {s.name} | {s.footprint_sf:,.0f} | {s.floors:g} | {s.gross_sf:,.0f} | {s.year_built} | {s.source} |")
    if job.to_remain:
        lines += ["", "**To remain / protect:** " + "; ".join(job.to_remain)]
    lines += ["", "## Biggest risks / gaps", ""] + [f"- {r}" for r in job.risks or ["(none listed)"]]
    lines += ["", "## Open questions for the client", ""] + [f"- **{i}** ({p}): {d}" for i, d, p in job.questions]
    ph = [k for k, v in rates.items() if v.placeholder]
    lines += ["", "## Placeholder rates used", "",
              f"{len(ph)} of {len(rates)} rates are still placeholders – replace them in Ahrens Rate Card.xlsx.",
              "", "## Sources", ""] + [f"- {s}" for s in job.sources]
    return "\n".join(lines) + "\n"


def cmd_build(job_json: Path, root: Path, rates_path: Path) -> Path:
    job = Job.from_dict(json.loads(Path(job_json).read_text()))
    rates = load_rates(rates_path)
    crew, persf = price_crew_time(job, rates), price_per_sf(job, rates)
    folder = job_folder(root, job.date, job.client, job.project)
    stem = safe_name(job.project)
    wb = folder / f"{stem} - Takeoff.xlsx"
    write_workbook(job, rates, wb)
    errs = recalc(wb)
    if errs:
        raise SystemExit(f"Workbook formula errors: {errs}")
    render_letter(job, crew.total, folder)
    (folder / f"{stem} - Bid Intel.md").write_text(bid_intel(job, rates, crew, persf), encoding="utf-8")
    (folder / "job.json").write_text(json.dumps(job.to_dict(), indent=1))
    append_job(root / "Bid Log.xlsx", job, crew, persf, folder)
    print(folder)
    return folder


def cmd_stop(reason: str, root: Path, client: str, project: str) -> Path:
    folder = job_folder(root, date.today(), client, project)
    note = folder / f"STOPPED – {safe_name(reason)}.md"
    note.write_text(f"# Stopped: {reason}\n\nNo bid was prepared. See the request for what was supplied.\n",
                    encoding="utf-8")
    print(folder)
    return folder


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="bidpkg")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("init"); p.add_argument("--root", type=Path, required=True)
    p = sub.add_parser("build"); p.add_argument("job_json", type=Path)
    p.add_argument("--root", type=Path, required=True); p.add_argument("--rates", type=Path, required=True)
    p = sub.add_parser("stop"); p.add_argument("reason"); p.add_argument("--root", type=Path, required=True)
    p.add_argument("--client", default="Unknown"); p.add_argument("--project", default="Unidentified")
    p = sub.add_parser("calibrate"); p.add_argument("csv", type=Path)
    p.add_argument("--rates", type=Path, required=True); p.add_argument("--out", type=Path, required=True)
    a = ap.parse_args(argv)
    if a.cmd == "init":
        cmd_init(a.root)
    elif a.cmd == "build":
        cmd_build(a.job_json, a.root, a.rates)
    elif a.cmd == "stop":
        cmd_stop(a.reason, a.root, a.client, a.project)
    elif a.cmd == "calibrate":
        jobs = cal.load_past_jobs(a.csv)
        res = cal.calibrate(jobs)
        cal.write_calibration(res, jobs, a.out)
        changed = cal.apply_to_rate_card(res, a.rates)
        print(json.dumps({"changed": changed, "derived": res}, indent=1, default=str))
    return 0


if __name__ == "__main__":
    sys.exit(main())
