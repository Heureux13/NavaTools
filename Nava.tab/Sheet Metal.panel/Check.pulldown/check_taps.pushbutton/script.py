# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from pyrevit import revit, script
from Autodesk.Revit.DB import BuiltInCategory, FabricationPart, FilteredElementCollector
from revit_duct import RevitDuct
from size import Size

# Button display information
# =================================================
__title__ = "Check Taps"
__doc__ = """
Lists all visible fabrication taps whose Connector0 value is not in tap_values.
"""


# Parameters to check for taps (case-insensitive)
# Edit these lists as needed.
tap_values = [
    "tap",
]

tap_families = [
    "boot tap",
    "straight tap",
]

target_families = set((x or "").strip().lower() for x in tap_families if (x or "").strip())


def _normalize(value):
    return (value or "").strip().lower()


def _is_tap_family(duct):
    family = _normalize(duct.family)
    return family in target_families


def _is_round_tap(duct):
    size_value = (duct.size or "").strip()
    if not size_value:
        return False

    parsed_size = Size(size_value)
    in_shape = parsed_size.in_shape()
    out_shape = parsed_size.out_shape()

    # Treat as round only when both sides are round.
    return in_shape == "round" and out_shape == "round"


doc = revit.doc
view = revit.active_view
output = script.get_output()

fab_parts = (
    FilteredElementCollector(doc, view.Id)
    .OfClass(FabricationPart)
    .OfCategory(BuiltInCategory.OST_FabricationDuctwork)
    .WhereElementIsNotElementType()
    .ToElements()
)

if not fab_parts:
    output.print_md("## No fabrication ductwork found in this view.")
    script.exit()

allowed_conn0_values = set(_normalize(x) for x in tap_values if _normalize(x))
if not allowed_conn0_values:
    output.print_md("## tap_values is empty. Add at least one allowed Connector0 value.")
    script.exit()

failures = []
for part in fab_parts:
    duct = RevitDuct(doc, view, part)

    if not _is_tap_family(duct):
        continue

    # Only report taps that are not round.
    if _is_round_tap(duct):
        continue

    conn0 = (duct.connector_0_type or "").strip() or "(empty)"
    if _normalize(conn0) in allowed_conn0_values:
        continue

    conn1 = (duct.connector_1_type or "").strip() or "(empty)"
    conn2 = (duct.connector_2_type or "").strip() or "(empty)"
    family = (duct.family or "").strip() or "(No Family)"
    service = (duct.service or "").strip() or "(No Service)"
    failures.append({
        "id": part.Id,
        "family": family,
        "service": service,
        "conn0": conn0,
        "conn1": conn1,
        "conn2": conn2,
    })

if not failures:
    output.print_md("## All checked taps have allowed Connector0 values.")
    script.exit()

output.print_md("## Taps with Connector0 not in tap_values")
output.print_md("**Count:** {}".format(len(failures)))

for item in sorted(failures, key=lambda x: (x["service"].lower(), x["family"].lower(), x["id"].IntegerValue)):
    output.print_md(
        "- ID: {} | Family: {} | Service: {} | C0: {} | C1: {} | C2: {}".format(
            output.linkify(item["id"]),
            item["family"],
            item["service"],
            item["conn0"],
            item["conn1"],
            item["conn2"],
        )
    )
