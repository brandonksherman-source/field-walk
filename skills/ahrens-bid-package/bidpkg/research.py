"""Lookups that work from a cloud session with no browser: geocoding, OSM footprints, St. Louis parcels, photo GPS."""
from __future__ import annotations
import json
import math
import time
import urllib.parse
import urllib.request
from pathlib import Path

from PIL import Image

UA = "ahrens-bid-research/0.1"
R = 6378137.0
SQM_TO_SQFT = 10.7639
STL_PARCELS = "https://maps6.stlouis-mo.gov/arcgis/rest/services/St_Louis_Parcels/MapServer/0/query"
OSM = "https://api.openstreetmap.org/api/0.6"
PARCEL_FIELDS = ("SITEADDR", "OWNERNAME", "AsrLandUse1", "FirstYearBuilt", "LastYearBuilt", "NbrOfBldgsCom",
                 "NbrOfBldgsRes", "SQFT", "CityBlock", "HANDLE")
_last_nominatim = 0.0


def _get_json(url: str):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    for attempt in (1, 2):
        try:
            with urllib.request.urlopen(req, timeout=20) as r:
                return json.loads(r.read().decode())
        except Exception:
            if attempt == 2:
                raise
            time.sleep(2)


def polygon_area_sf(coords: list[tuple[float, float]]) -> float:
    """Area of a (lon, lat) ring in square feet (equirectangular, fine at building scale)."""
    if len(coords) < 3:
        return 0.0
    lat0 = math.radians(coords[0][1])
    pts = [(math.radians(lon) * R * math.cos(lat0), math.radians(lat) * R) for lon, lat in coords]
    a = sum(x1 * y2 - x2 * y1 for (x1, y1), (x2, y2) in zip(pts, pts[1:] + pts[:1]))
    return abs(a / 2) * SQM_TO_SQFT


def geocode(address: str) -> list[dict]:
    global _last_nominatim
    wait = 1.0 - (time.time() - _last_nominatim)
    if wait > 0:
        time.sleep(wait)
    _last_nominatim = time.time()
    url = "https://nominatim.openstreetmap.org/search?" + urllib.parse.urlencode(
        {"format": "json", "q": address, "limit": 5})
    return [{"lat": float(x["lat"]), "lon": float(x["lon"]), "display_name": x["display_name"]}
            for x in _get_json(url)]


def _ring(way: dict, nodes: dict) -> list[tuple[float, float]]:
    return [(nodes[n]["lon"], nodes[n]["lat"]) for n in way.get("nodes", []) if n in nodes]


def _element_area(el: dict, nodes: dict, ways: dict) -> float:
    if el["type"] == "way":
        return polygon_area_sf(_ring(el, nodes))
    total = 0.0
    for m in el.get("members", []):
        if m["type"] == "way" and m["ref"] in ways:
            a = polygon_area_sf(_ring(ways[m["ref"]], nodes))
            total += -a if m.get("role") == "inner" else a
    return total


def _index(elements):
    nodes = {e["id"]: e for e in elements if e["type"] == "node"}
    ways = {e["id"]: e for e in elements if e["type"] == "way"}
    return nodes, ways


def _complete(el, nodes, ways) -> bool:
    if el["type"] == "way":
        return all(n in nodes for n in el.get("nodes", []))
    return all(m["ref"] in ways and _complete(ways[m["ref"]], nodes, ways)
               for m in el.get("members", []) if m["type"] == "way")


def osm_buildings(lat: float, lon: float, half_box_deg: float = 0.002) -> list[dict]:
    bbox = f"{lon - half_box_deg},{lat - half_box_deg * 0.75},{lon + half_box_deg},{lat + half_box_deg * 0.75}"
    els = _get_json(f"{OSM}/map.json?bbox={bbox}")["elements"]
    nodes, ways = _index(els)
    out = []
    for el in els:
        tags = el.get("tags") or {}
        if el["type"] not in ("way", "relation") or "building" not in tags:
            continue
        if not _complete(el, nodes, ways):   # crosses the box edge: fetch the whole thing
            full = _get_json(f"{OSM}/{el['type']}/{el['id']}/full.json")["elements"]
            fn, fw = _index(full)
            area = _element_area(el, fn, fw)
        else:
            area = _element_area(el, nodes, ways)
        out.append({"osm_type": el["type"], "osm_id": el["id"], "name": tags.get("name", ""),
                    "building": tags.get("building", ""), "levels": tags.get("building:levels", ""),
                    "footprint_sf": round(area)})
    return sorted(out, key=lambda b: -b["footprint_sf"])


def stl_parcels(lat: float, lon: float, radius_m: float = 25) -> list[dict]:
    """City of St. Louis assessor parcels within radius_m of the point; [] outside the city."""
    q = {"f": "json", "geometry": f"{lon},{lat}", "geometryType": "esriGeometryPoint", "inSR": 4326,
         "spatialRel": "esriSpatialRelIntersects", "distance": radius_m, "units": "esriSRUnit_Meter",
         "outFields": ",".join(PARCEL_FIELDS), "returnGeometry": "false"}
    d = _get_json(STL_PARCELS + "?" + urllib.parse.urlencode(q))
    return [{k: (v.strip() if isinstance(v, str) else v) for k, v in f["attributes"].items()}
            for f in d.get("features", [])]


def _deg(dms, ref) -> float:
    d, m, s = (float(x) for x in dms)
    v = d + m / 60 + s / 3600
    return -v if ref in ("S", "W") else v


def photo_gps(path: Path) -> tuple[float, float] | None:
    try:
        gps = Image.open(path).getexif().get_ifd(0x8825)
    except Exception:
        return None
    if not gps or 2 not in gps or 4 not in gps:
        return None
    return _deg(gps[2], gps.get(1, "N")), _deg(gps[4], gps.get(3, "E"))
