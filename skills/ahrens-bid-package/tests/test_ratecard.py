import openpyxl, pytest
from bidpkg.ratecard import create_rate_card, load_rates, RateCardError

def test_rate_card_round_trip(tmp_path):
    p = tmp_path/"Ahrens Rate Card.xlsx"; create_rate_card(p)
    r = load_rates(p)
    assert r["mech_demo_sf"].value == 3.50 and r["mech_demo_sf"].placeholder
    assert r["bond_pct"].value == 0.015
    assert r["mech_demo_sf"].source == "market placeholder – replace with Ahrens rates/quotes"

def test_placeholder_cells_are_blue_on_yellow(tmp_path):
    p = tmp_path/"rc.xlsx"; create_rate_card(p)
    ws = openpyxl.load_workbook(p)["Rates"]
    c = ws.cell(5, 4)
    assert c.font.color.rgb.endswith("0000FF") and c.fill.fgColor.rgb.endswith("FFFF00")

def test_blank_rate_names_the_cell(tmp_path):
    p = tmp_path/"rc.xlsx"; create_rate_card(p)
    wb = openpyxl.load_workbook(p); ws = wb["Rates"]
    row = next(r for r in range(5, ws.max_row+1) if ws.cell(r,1).value == "haul_load")
    ws.cell(row,4).value = None; wb.save(p)
    with pytest.raises(RateCardError, match=rf"Rates!D{row} \(haul_load\)"):
        load_rates(p)

def test_text_rate_names_the_cell(tmp_path):
    p = tmp_path/"rc.xlsx"; create_rate_card(p)
    wb = openpyxl.load_workbook(p); ws = wb["Rates"]
    ws.cell(5,4).value = "call Bob"; wb.save(p)
    with pytest.raises(RateCardError, match=r"Rates!D5"):
        load_rates(p)
