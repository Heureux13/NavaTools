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
from Autodesk.Revit.DB import *

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

hanger_parameters = {
    '_umi_duct_supporting_weight',
    'mark',
}

duct_parameters = {
    '_umi_duc_run_weight',
    'mark'
}


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
    total_length = sum(d.length or 0 for d in selected_ducts)
    total_weight = sum(safe_float(d.weight) or 0 for d in selected_ducts)

    # Get hangers that intersect with selected ducts via bounding box
    hangers = set()  # Use set to avoid duplicates

    for duct in selected_ducts:
        bbox = duct.element.get_BoundingBox(None)
        if bbox:
            # Create filter for elements intersecting this bounding box
            outline = Outline(bbox.Min, bbox.Max)
            bbox_filter = BoundingBoxIntersectsFilter(outline)

            # Collect hangers intersecting this duct
            intersecting_hangers = FilteredElementCollector(doc)\
                .OfCategory(BuiltInCategory.OST_FabricationHangers)\
                .WherePasses(bbox_filter)\
                .WhereElementIsNotElementType()\
                .ToElements()

            for h in intersecting_hangers:
                hangers.add(h)

    hangers = list(hangers)  # Convert back to list
    duct_size = selected_ducts[0].size

    # Display results and set hanger marks
    if hangers:
        weight_per_hanger = total_weight / len(hangers)
        hanger_ids = [h.Id for h in hangers]

        RevitElement.select_many(uidoc, hangers)

        output.print_md("---")
        output.print_md("# Found {} hanger(s): {}".format(
            len(hangers),
            output.linkify(hanger_ids)
        ))

        # Set parameters on hangers and ducts
        with revit.Transaction("Set Hanger and Duct Marks"):
            for i, hanger in enumerate(hangers, start=1):
                output.print_md("### {} | ID: {} | Supporting: {:.2f} lbs".format(
                    i,
                    output.linkify(hanger.Id),
                    weight_per_hanger
                ))

                # Try each hanger parameter in order, set the first writable one
                set_success = False
                for parameter_name in hanger_parameters:
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

            # Set run weight on each selected duct
            for d in selected_ducts:
                set_parameter = None
                for parameter_name in duct_parameters:
                    p = d.element.LookupParameter(parameter_name)
                    if not p:
                        continue
                    elif p.IsReadOnly:
                        continue
                    else:
                        set_parameter = p
                        break

                if set_parameter:
                    set_parameter.Set(total_weight)

        output.print_md("---")

    # Display duct information
    duct_element_ids = [d.element.Id for d in selected_ducts]
    total_length_ft = total_length / 12.0 if total_length else 0.0
    lbs_per_ft = (total_weight / total_length_ft) if total_length_ft else 0.0
    output.print_md("# Selected Ducts Information")
    output.print_md("## Qty: {} | Size: {} | Length: {:06.2f} ft | Weight: {:06.2f} lbs | Weight/ft: {:06.2f} lbs/ft | {}".format(
        len(selected_ducts),
        duct_size,
        total_length_ft,
        total_weight,
        lbs_per_ft,
        output.linkify(duct_element_ids)
    ))
