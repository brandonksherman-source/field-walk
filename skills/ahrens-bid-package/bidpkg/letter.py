"""The Ahrens proposal letter, laid out like the Building 121/123 proposal (Oct 12, 2026).

No signature and no person's name: the sender signs at review.
"""
from __future__ import annotations
import subprocess
import tempfile
from pathlib import Path

import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.shared import Pt, Inches, RGBColor

from .models import Job
from .paths import safe_name
from .scope import included_bullets, exclusions

LOGO = Path(__file__).resolve().parents[1] / "templates" / "letterhead_logo.png"
LETTERHEAD = ["Ahrens Contracting, Inc.", "140 Lafayette Ave.", "St. Louis, MO 63104",
              "Ph 314-631-7799   Fax 314-621-0977", "WBE, WBENC, Avetta & ISN Certified"]
FOOTER = "Ahrens Contracting Inc. • 140 Lafayette Avenue • St. Louis, MO 63104"
OPENING = "Ahrens Contracting will provide the labor, material and equipment to complete the following scope of work."
CONTACT = "If you have any questions, please contact us at 314-631-7799."
DISCLAIMER = ("This proposal is contingent upon and shall not be binding or valid unless and until acceptable "
              "subcontract terms are agreed to in writing by Ahrens Contracting, Inc. (which acceptance may be "
              "withheld for any reason and the submission of this proposal is not an acceptance of any tendered or "
              "proposed subcontract terms); including, but not limited to, terms with respect to payment, schedule "
              "for the work and disputes.  The party receiving this proposal may not rely upon this proposal unless "
              "and until this contingency is satisfied.")
DRAFT = "DRAFT – for review"
FONT = "Calibri"
HONORIFICS = {"mr", "mrs", "ms", "miss", "dr", "mx", "prof"}


def _para(container, text="", bold=False, underline=False, size=11, align=None, space_after=8, color=None,
          highlight=False):
    p = container.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    if align:
        p.alignment = align
    if text:
        r = p.add_run(text)
        r.bold, r.underline = bold, underline
        r.font.size, r.font.name = Pt(size), FONT
        if color:
            r.font.color.rgb = color
        if highlight:
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    return p


def _mixed(container, parts, space_after=8):
    """parts: list of (text, highlight)"""
    p = container.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    for text, hl in parts:
        r = p.add_run(text)
        r.font.size, r.font.name = Pt(11), FONT
        if hl:
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    return p


def _bullets(doc, items):
    for it in items:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.left_indent = Inches(0.75)
        r = p.add_run(it)
        r.font.size, r.font.name = Pt(11), FONT


def _letterhead(section):
    hdr = section.header
    hdr.is_linked_to_previous = False
    for p in list(hdr.paragraphs):
        p._element.getparent().remove(p._element)
    t = hdr.add_table(rows=1, cols=3, width=Inches(6.5))
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    t.autofit = False
    widths = (Inches(1.1), Inches(4.2), Inches(1.2))
    for col, w in zip(t.columns, widths):
        col.width = w
        for c in col.cells:
            c.width = w
    logo_cell, text_cell, draft_cell = t.rows[0].cells
    lp = logo_cell.paragraphs[0]
    lp.add_run().add_picture(str(LOGO), height=Inches(0.95))
    for i, line in enumerate(LETTERHEAD):
        p = text_cell.paragraphs[0] if i == 0 else text_cell.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(line)
        r.font.name, r.font.size, r.bold = "Arial", Pt(12 if i == 0 else 11), i == 0
    dp = draft_cell.paragraphs[0]
    dp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = dp.add_run(DRAFT)
    r.bold, r.font.size, r.font.name = True, Pt(9), "Arial"
    r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    hdr.add_paragraph().paragraph_format.space_after = Pt(6)

    ftr = section.footer
    fp = ftr.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = fp.add_run(FOOTER)
    r.font.name, r.font.size, r.bold = "Arial", Pt(10), True
    r.font.color.rgb = RGBColor(0xA6, 0xA6, 0xA6)


def build_docx(job: Job, lump_sum: float, path: Path) -> Path:
    doc = docx.Document()
    st = doc.styles["Normal"]
    st.font.name, st.font.size = FONT, Pt(11)
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.5), Inches(11)
    sec.left_margin = sec.right_margin = Inches(1)
    sec.top_margin, sec.bottom_margin = Inches(0.6), Inches(0.7)
    sec.header_distance, sec.footer_distance = Inches(0.4), Inches(0.4)
    _letterhead(sec)

    rec = (job.recipient or "").strip() or None
    names = [w for w in (rec or "").split() if w.rstrip(".").lower() not in HONORIFICS]
    first = names[0] if names else None
    _para(doc, f"{job.date:%B} {job.date.day}, {job.date.year}", space_after=14)
    _mixed(doc, [("To: ", False), (rec or "[RECIPIENT]", rec is None)], space_after=14)
    _para(doc, f"RE: {job.project} Demo", space_after=14)
    _mixed(doc, [("Dear ", False), (first or "[FIRST NAME]", first is None), (",", False)], space_after=14)
    _para(doc, OPENING, space_after=14)
    _para(doc, "Items Included:", underline=True, space_after=6)
    _bullets(doc, included_bullets(job.switches) + list(job.included_bullets_extra))
    _para(doc, "", space_after=4)
    _para(doc, f"LUMP SUM: ${lump_sum:,.2f}", bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=14)
    _para(doc, "Exclusions:", bold=True, space_after=6)
    _bullets(doc, exclusions(job.switches))
    _para(doc, "", space_after=10)
    _para(doc, CONTACT, space_after=18)
    _para(doc, "Best Regards,", space_after=48)   # space for the sender to sign and print their name
    _para(doc, "Ahrens Contracting, Inc.", space_after=14)
    _para(doc, DISCLAIMER, size=8, space_after=0)
    doc.save(path)
    return path


def to_pdf(docx_path: Path) -> Path:
    docx_path = Path(docx_path)
    with tempfile.TemporaryDirectory() as td:
        subprocess.run(["soffice", f"-env:UserInstallation=file://{Path(td) / 'profile'}", "--headless",
                        "--convert-to", "pdf", "--outdir", str(docx_path.parent), str(docx_path)],
                       check=True, capture_output=True, timeout=180)
    return docx_path.with_suffix(".pdf")


def render_letter(job: Job, lump_sum: float, out_dir: Path) -> tuple[Path, Path]:
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    d = build_docx(job, lump_sum, out_dir / f"{safe_name(job.project)} - Proposal.docx")
    return d, to_pdf(d)
