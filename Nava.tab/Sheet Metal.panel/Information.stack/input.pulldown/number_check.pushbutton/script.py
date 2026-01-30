# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, ElementId, VisibleInViewFilter
from pyrevit import revit, script
from System.Collections.Generic import List
from revit_duct import RevitDuct


# Button info
# ===================================================
__title__ = "Check Fab Numbers"
__doc__ = """
Select ducts where Fabrication Notes and Item Number match
"""

# Variables
# ==================================================
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
view = revit.active_view
output = script.get_output()

# Helpers
# ========================================================================


def get_param_value(param):
    try:
        if param.StorageType == 0:  # None
            return None
        if param.AsString():
            return param.AsString()
        if param.AsValueString():
            return param.AsValueString()
        if param.StorageType == 1:  # Double
            return param.AsDouble()
        if param.StorageType == 2:  # Integer
            return param.AsInteger()
        if param.StorageType == 3:  # ElementId
            return param.AsElementId().IntegerValue
    except Exception:
        return None


# Main Code
# ==================================================
try:
    # Collect only fabrication ductwork strictly visible in the active view
    fab_duct = (FilteredElementCollector(doc, view.Id)
                .OfCategory(BuiltInCategory.OST_FabricationDuctwork)
                .WhereElementIsNotElementType()
                .WherePasses(VisibleInViewFilter(doc, view.Id))
                .ToElements())

    all_duct = list(fab_duct)
    if not all_duct:
        output.print_md("## No fabrication ducts found in view")
        script.exit()

    # Build parameter -> value -> elements map
    # Key is "Fabrication Notes | Item Number" combination
    param_groups = {}
    for d in all_duct:
        fab_notes = None
        item_number = None

        fn_param = d.LookupParameter("Fabrication Notes")
        in_param = d.LookupParameter("Item Number")
        if fn_param:
            fab_notes = get_param_value(fn_param)
        if in_param:
            item_number = get_param_value(in_param)

        # Skip if either is empty
        empty_vals = {None, "", "**"}
        if fab_notes in empty_vals or item_number in empty_vals:
            continue

        # Group by combination of both parameters
        composite_key = "{} | {}".format(fab_notes, item_number)
        if composite_key not in param_groups:
            param_groups[composite_key] = []
        param_groups[composite_key].append(d)

    if not param_groups:
        output.print_md(
            "## No ducts found with Fabrication Notes and Item Number populated")
        script.exit()

    # Select ducts that have matching combinations (duplicates)
    duct_run = []
    for key, ducts in param_groups.items():
        if len(ducts) > 1:
            duct_run.extend(ducts)

    if not duct_run:
        output.print_md("## No matching ducts found")
        script.exit()

    # Select ducts in Revit
    duct_ids = List[ElementId]()
    for d in duct_run:
        duct_ids.Add(d.Id)
    uidoc.Selection.SetElementIds(duct_ids)

    output.print_md("## Selected {} ducts".format(len(duct_run)))
    output.print_md("---")

    fil_ducts = []
    for d in duct_run:
        try:
            fil_ducts.append(RevitDuct(doc, view, d))
        except Exception:
            pass

    for i, fil in enumerate(fil_ducts, start=1):
        try:
            fn_param = fil.element.LookupParameter("Fabrication Notes")
            in_param = fil.element.LookupParameter("Item Number")
            fab_notes = get_param_value(fn_param) if fn_param else ""
            item_number = get_param_value(in_param) if in_param else ""
        except Exception:
            fab_notes = ""
            item_number = ""
        output.print_md(
            '### No: {:03} | ID: {} | Notes: {} | Item: {} | Length: {:06.2f}" | Size: {}'.format(
                i,
                output.linkify(fil.element.Id),
                fab_notes,
                item_number,
                fil.length if fil.length is not None else 0.0,
                fil.size,
            )
        )

    element_ids = [d.element.Id for d in fil_ducts]
    output.print_md(
        "# Total elements {}, {}".format(
            len(fil_ducts),
            output.linkify(element_ids)
        )
    )

except Exception as e:
    output.print_md("## Error: {}".format(str(e)))
