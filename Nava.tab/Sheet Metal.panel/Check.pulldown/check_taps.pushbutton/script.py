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

# Button display information
# =================================================
__title__ = "Check Taps"
__doc__ = """
Lists all visible fabrication taps whose Connector0 end condition is not Tap.
"""


def _param_text(element, param_names):
    for name in param_names:
        p = element.LookupParameter(name)
        if not p:
            continue
        value = p.AsString() or p.AsValueString() or ""
        value = value.strip()
        if value:
            return value
    return ""


def _is_tap_family(element):
    family = _param_text(element, ["Family", "Family Name"]).lower()
    return "tap" in family


def _service_name(element):
    return _param_text(
        element,
        ["Fabrication Service Name", "Fabrication Service", "Service"]
    ) or "(No Service)"


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

failures = []
for part in fab_parts:
    if not _is_tap_family(part):
        continue

    conn0 = _param_text(part, ["NaviateDBS_Connector0_EndCondition", "Connector0_EndCondition"]) or "(empty)"
    if conn0.lower() == "tap":
        continue

    conn1 = _param_text(part, ["NaviateDBS_Connector1_EndCondition", "Connector1_EndCondition"]) or "(empty)"
    conn2 = _param_text(part, ["NaviateDBS_Connector2_EndCondition", "Connector2_EndCondition"]) or "(empty)"
    family = _param_text(part, ["Family", "Family Name"]) or "(No Family)"
    failures.append({
        "id": part.Id,
        "family": family,
        "service": _service_name(part),
        "conn0": conn0,
        "conn1": conn1,
        "conn2": conn2,
    })

if not failures:
    output.print_md("## All visible taps have Connector0 = Tap.")
    script.exit()

output.print_md("## Taps with Connector0 != Tap")
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
