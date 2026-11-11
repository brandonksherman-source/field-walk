import math, pytest
from PIL import Image
from bidpkg.research import polygon_area_sf, photo_gps, geocode, osm_buildings, stl_parcels

R = 6378137

def square_ring(lat, lon, side_m):
    dlat = side_m / R * 180 / math.pi
    dlon = side_m / (R * math.cos(math.radians(lat))) * 180 / math.pi
    return [(lon, lat), (lon + dlon, lat), (lon + dlon, lat + dlat), (lon, lat + dlat), (lon, lat)]

def test_area_of_100ft_square():
    assert polygon_area_sf(square_ring(38.63, -90.21, 30.48)) == pytest.approx(10000, rel=0.01)

def test_area_degenerate_ring_is_zero():
    assert polygon_area_sf([(0, 0), (1, 1)]) == 0

def _dms(x):
    x = abs(x); d = int(x); m = int((x - d) * 60); s = round(((x - d) * 60 - m) * 60 * 100)
    return ((d, 1), (m, 1), (s, 100))

def make_gps_jpeg(tmp_path, lat, lon):
    im = Image.new("RGB", (10, 10)); ex = Image.Exif()
    gps = {1: "N" if lat >= 0 else "S", 2: tuple(a / b for a, b in _dms(lat)),
           3: "E" if lon >= 0 else "W", 4: tuple(a / b for a, b in _dms(lon))}
    ex[0x8825] = gps
    p = tmp_path / "g.jpg"; im.save(p, exif=ex); return p

def test_photo_without_gps(tmp_path):
    p = tmp_path / "x.jpg"; Image.new("RGB", (10, 10)).save(p)
    assert photo_gps(p) is None

def test_photo_with_gps(tmp_path):
    p = make_gps_jpeg(tmp_path, 38.6353, -90.2205)
    assert photo_gps(p) == pytest.approx((38.6353, -90.2205), abs=1e-4)

@pytest.mark.network
def test_wells_fargo_building_e_live():
    g = geocode("2700 Pine St, St. Louis, MO")
    lat, lon = float(g[0]["lat"]), float(g[0]["lon"])
    assert any((p.get("OWNERNAME") or "").upper().startswith("A G EDWARDS") for p in stl_parcels(lat, lon))
    assert max(b["footprint_sf"] for b in osm_buildings(lat, lon)) > 40000

@pytest.mark.network
def test_nonsense_address_live():
    assert geocode("zzzz no such place 99999") == []

from bidpkg.research import clipped_area_sf, osm_polygon, stl_parcel_polygon

def test_clipped_area_half_square():
    sq = square_ring(38.63, -90.21, 30.48)                     # 100 ft square
    (x0, y0), (x1, _), _, (_, y1), _ = sq
    left_half = [(x0, y0), ((x0 + x1) / 2, y0), ((x0 + x1) / 2, y1), (x0, y1), (x0, y0)]
    assert clipped_area_sf([sq], [left_half]) == pytest.approx(5000, rel=0.01)

def test_clipped_area_no_overlap_is_zero():
    a = square_ring(38.63, -90.21, 30.48); b = square_ring(38.64, -90.21, 30.48)
    assert clipped_area_sf([a], [b]) == 0

@pytest.mark.network
def test_wells_fargo_campus_clipped_to_building_e_parcel_live():
    area = clipped_area_sf(osm_polygon("relation", 19553658), stl_parcel_polygon("10987000010"))
    assert area == pytest.approx(56749, rel=0.10)

from bidpkg.research import stacked_gross_sf

def test_stacked_gross_nested_tiers_not_double_counted():
    big = square_ring(38.63, -90.21, 30.48)          # 10,000 SF, 2 floors
    (x0, y0), (x1, _), _, (_, y1), _ = big
    small = [(x0, y0), ((x0 + x1) / 2, y0), ((x0 + x1) / 2, y1), (x0, y1), (x0, y0)]   # 5,000 SF inside it, 4 floors
    g = stacked_gross_sf([([big], 2, 0), ([small], 4, 0)])
    assert g["gross_sf"] == pytest.approx(10000 * 2 + 5000 * 2, rel=0.01)
    assert g["footprint_sf"] == pytest.approx(10000, rel=0.01) and g["max_floors"] == 4
