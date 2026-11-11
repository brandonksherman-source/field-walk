import docx
from bidpkg.letter import render_letter

def all_text(p):
    d = docx.Document(p)
    parts = [x.text for x in d.paragraphs]
    for s in d.sections:
        for part in (s.header, s.footer, s.first_page_header):
            parts += [x.text for x in part.paragraphs]
            for t in part.tables:
                for row in t.rows:
                    for c in row.cells:
                        parts += [x.text for x in c.paragraphs]
    return "\n".join(parts)

def test_letter_content(tmp_path, job):
    docx_p, pdf_p = render_letter(job, 73777, tmp_path)
    assert docx_p.name == "Fixture Bldg - Proposal.docx" and pdf_p.name == "Fixture Bldg - Proposal.pdf"
    text = all_text(docx_p)
    assert "LUMP SUM: $73,777.00" in text
    assert "Items Included:" in text and "Exclusions:" in text
    assert "WBE, WBENC, Avetta & ISN Certified" in text and "Ph 314-631-7799" in text
    assert "Ahrens Contracting Inc. • 140 Lafayette Avenue • St. Louis, MO 63104" in text
    assert "Ahrens Contracting will provide the labor, material and equipment to complete the following scope of work." in text
    assert "To: Bruce Byrd" in text and "Dear Bruce," in text and "RE: Fixture Bldg Demo" in text
    assert "November 11, 2026" in text
    assert "Roger Kent" not in text and "VP of Operations" not in text
    assert "DRAFT – for review" in text
    assert "This proposal is contingent upon" in text
    assert pdf_p.exists() and pdf_p.stat().st_size > 5000

def test_switched_off_line_is_excluded_not_included(tmp_path, job):
    job.switches["backfill"] = False
    text = all_text(render_letter(job, 1000, tmp_path)[0])
    inc, exc = text.split("Exclusions:")
    assert "Backfill basements" not in inc and "\nBackfill\n" in exc + "\n"

def test_missing_recipient_placeholder(tmp_path, job):
    job.recipient = None
    text = all_text(render_letter(job, 1000, tmp_path)[0])
    assert "To: [RECIPIENT]" in text and "Dear [FIRST NAME]," in text

def test_project_with_slash_gets_safe_filename(tmp_path, job):
    job.project = "Building 121/123"
    d, p = render_letter(job, 1000, tmp_path)
    assert d.name == "Building 121-123 - Proposal.docx" and "RE: Building 121/123 Demo" in all_text(d)
