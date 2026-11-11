"""Every scope line Ahrens bids, with its letter wording when included or excluded.

Exclusion wording follows the Building 121/123 proposal (Oct 12, 2026).
"""
from __future__ import annotations
from dataclasses import dataclass


@dataclass(frozen=True)
class ScopeLine:
    key: str
    label: str
    included_bullet: str
    exclusion_text: str


_LINES = [
    ScopeLine("mobilization", "Mobilization",
              "Mobilization and demobilization of all labor and equipment",
              "Mobilization and demobilization"),
    ScopeLine("permits", "Demolition permits",
              "Demolition permit(s) for our work",
              "All permits (by owner)"),
    ScopeLine("fencing", "Fencing & protection",
              "Site fencing and protection during demolition",
              "Site fencing and/or security"),
    ScopeLine("utilities", "Utility disconnects / cut & cap",
              "Disconnect, cut and cap all utilities (electric, gas, water, sewer)",
              "All electrical or mechanical isolations or re-routes"),
    ScopeLine("tap_destroys", "Tap destroys",
              "All tap destroys",
              "All tap destroys (if required)"),
    ScopeLine("hazmat", "Hazmat survey & abatement",
              "Hazardous materials survey, asbestos abatement and removal of universal waste "
              "(bulbs/ballasts, Freon, household hazardous waste)",
              "All environmental work, including, but not limited to: asbestos abatement, Freon evacuation, "
              "bulbs/ballasts, household hazardous waste, environmental surveys, hazardous and special waste"),
    ScopeLine("demolition", "Demolition & loading",
              "Complete demolition of all structures to slab or existing ground/street elevations; "
              "removal and loading of all masonry, brick and CMU block; remove all steel framework",
              "Building demolition"),
    ScopeLine("haul", "Trucking / haul",
              "Haul all debris from site",
              "Hauling and trucks for debris and salvage (paid by owner)"),
    ScopeLine("disposal", "Tipping / disposal",
              "All tipping and disposal fees",
              "Tipping or disposal fees (paid by owner)"),
    ScopeLine("scrap_credit", "Scrap / salvage credit",
              "All salvage and scrap metal retained by Ahrens Contracting (credited in this price)",
              "All salvage generated (paid to owner)"),
    ScopeLine("backfill", "Backfill & compaction",
              "Backfill basements and voids with clean fill, compacted",
              "Backfill"),
    ScopeLine("seed_straw", "Seed / straw & restoration",
              "Final grade, seed and straw disturbed areas",
              "Seed/straw & site restoration"),
    ScopeLine("trees", "Tree & stump removal",
              "Removal of trees and stumps within the work area",
              "Removal of trees and/or stumps"),
    ScopeLine("swppp", "Erosion control / SWPPP",
              "Erosion control and SWPPP measures during demolition",
              "Erosion control and SWPPP measures"),
]

SCOPE_LINES: dict[str, ScopeLine] = {l.key: l for l in _LINES}

FIXED_EXCLUSIONS: list[str] = [
    "SWPP maintenance and reports after demolition is complete",
    "Pavement repair or damage",
    "Permitting beyond our work",
    'Additional cost for "wet demolition"',
    "All work to be completed during 7-3:30. If work requires labor and/or equipment beyond these hours, "
    "an additional charge will be incurred.",
]


def default_switches() -> dict[str, bool]:
    return {k: True for k in SCOPE_LINES}


def included_bullets(switches: dict[str, bool]) -> list[str]:
    return [l.included_bullet for k, l in SCOPE_LINES.items() if switches.get(k, True)]


def exclusions(switches: dict[str, bool]) -> list[str]:
    return [l.exclusion_text for k, l in SCOPE_LINES.items() if not switches.get(k, True)] + FIXED_EXCLUSIONS
