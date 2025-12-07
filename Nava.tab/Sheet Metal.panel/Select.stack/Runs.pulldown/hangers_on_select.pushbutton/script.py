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
__doc__ = """
Total weight of selected duct / amount of hangers
"""

# Variables
# ==================================================
doc = revit.doc
uidoc = __revit__.ActiveUIDocument
view = revit.active_view
output = script.get_output()

print_parameter = [
    '_weight_supporting',
    'mark',
]


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
    selected_duct_ids = set(d.element.Id for d in selected_ducts)

    for hanger in all_hangers:
        # Print Primary Element parameter for diagnostics
        ref_param = hanger.LookupParameter("Primary Element")
        ref_val = ref_param.AsString() if ref_param else None
        output.print_md(
            "Hanger ID {} | Primary Element: {}".format(
                hanger.Id,
                ref_val
            ))

        size_param = hanger.LookupParameter("Size of Primary End")
        if size_param:
            hanger_size = size_param.AsString()
            # Match both normalized and original formats
            if (hanger_size == duct_size or normalize_size(hanger_size) == duct_size_normalized):
                # Check if hanger is connected to any selected duct
                # For FabricationPart, try using the 'Primary Element' reference
                ref_param = hanger.LookupParameter("Primary Element")
                if ref_param:
                    ref_id_str = ref_param.AsString()
                    try:
                        ref_id = int(ref_id_str)
                        if ref_id in selected_duct_ids:
                            hangers.append(hanger)
                    except Exception:
                        pass

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

                # Try each parameter in order, set the first writable one
                set_success = False
                for parameter_name in print_parameter:
                    param = hanger.LookupParameter(parameter_name)
                    if param and not param.IsReadOnly:
                        param.Set(weight_per_hanger)
                        set_success = True
                        break
                if not set_success:
                    output.print_md(
                        'Could not set any parameter on hanger ID {}'.format(
                            output.linkify(hanger.Id)
                        ))

        output.print_md("---")

    # Display duct information
    duct_element_ids = [d.element.Id for d in selected_ducts]
    output.print_md("# Selected Ducts Information")
    output.print_md("## Qty: {} | Size: {} | Length: {:06.2f} ft | Weight: {:06.2f} lbs | Weight/ft: {:06.2f} lbs/ft | {}".format(
        len(selected_ducts),
        duct_size,
        total_length / 12,
        total_weight,
        total_weight / (total_length / 12) if total_length > 0 else 0,
        output.linkify(duct_element_ids)
    ))
