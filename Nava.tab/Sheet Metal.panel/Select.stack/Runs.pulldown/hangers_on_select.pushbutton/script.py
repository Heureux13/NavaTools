# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from revit_element import RevitElement
from revit_duct import RevitDuct
from pyrevit import revit, script
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory

# Button info
# ===================================================
__title__ = "Hangers on Select"
__doc__ = """Calculate and display weight distribution on hangers for selected duct"""

# Variables
# ==================================================
doc = revit.doc
uidoc = __revit__.ActiveUIDocument
view = revit.active_view
output = script.get_output()


def safe_float(val):
    """Convert value to float, return 0.0 if conversion fails"""
    try:
        return float(val)
    except Exception:
        return 0.0


def normalize_size(size_string):
    """Remove 'x' and 'ø' symbols from size string for comparison"""
    if not size_string:
        return None
    return size_string.replace("ø", "").replace("x", "")


# Main Code
# ==================================================

# Get selected ducts (can be multiple)
selected_ducts = RevitDuct.from_selection(uidoc, doc, view)

if not selected_ducts:
    output.print_md("## Select one or more ducts first")
else:
    # Calculate total properties from all selected ducts
    total_length = sum(safe_float(RevitDuct.parse_length_string(
        d.centerline_length)) or 0 for d in selected_ducts)
    total_weight = sum(safe_float(d.weight) or 0 for d in selected_ducts)
    # Use first duct's size for hanger matching
    duct_size = selected_ducts[0].size

    # Get all fabrication hangers in the view
    all_hangers = FilteredElementCollector(doc, view.Id)\
        .OfCategory(BuiltInCategory.OST_FabricationHangers)\
        .WhereElementIsNotElementType()\
        .ToElements()

    # Find hangers matching the duct size
    hangers = []
    duct_size_normalized = normalize_size(duct_size)

    for hanger in all_hangers:
        size_param = hanger.LookupParameter("Size of Primary End")
        if size_param:
            hanger_size = size_param.AsString()
            # Match both normalized and original formats
            if (hanger_size == duct_size or
                    normalize_size(hanger_size) == duct_size_normalized):
                hangers.append(hanger)

    # Display results and set hanger marks
    if hangers:
        weight_per_hanger = round(total_weight / len(hangers), 2)
        hanger_ids = [h.Id for h in hangers]

        RevitElement.select_many(uidoc, hangers)

        output.print_md("---")
        output.print_md("# Found {} hanger(s): {}".format(
            len(hangers),
            output.linkify(hanger_ids)
        ))

        # Set Mark parameter on each hanger
        with revit.Transaction("Set Hanger Marks"):
            for i, hanger in enumerate(hangers, start=1):
                output.print_md("### {} | ID: {} | Supporting: {:.2f} lbs".format(
                    i,
                    output.linkify(hanger.Id),
                    weight_per_hanger
                ))

                mark_param = hanger.LookupParameter("Mark")
                if mark_param:
                    mark_param.Set(str(weight_per_hanger))

        output.print_md("---")

    # Display duct information
    duct_element_ids = [d.element.Id for d in selected_ducts]
    output.print_md("# Selected Ducts Information")
    output.print_md("## Qty: {} | Size: {} | Length: {:.2f} ft | Weight: {:.2f} lbs | Weight/ft: {:.2f} lbs/ft | {}".format(
        len(selected_ducts),
        duct_size,
        total_length / 12,
        total_weight,
        total_weight / (total_length / 12) if total_length > 0 else 0,
        output.linkify(duct_element_ids)
    ))
