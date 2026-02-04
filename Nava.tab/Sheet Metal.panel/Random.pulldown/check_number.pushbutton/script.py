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


def lookup_parameter_case_insensitive(element, param_name):
    """Case-insensitive parameter lookup"""
    param_name_lower = param_name.strip().lower()
    for param in element.Parameters:
        if param.Definition.Name.strip().lower() == param_name_lower:
            return param
    return None


skip_values = {
    'skip',
    '0',
}

check_parameters = {
    # "fabrication notes",
    "item number",
}

# Main Code
# ==================================================
try:
    # Collect fabrication ductwork in the active view
    fab_duct = (FilteredElementCollector(doc, view.Id)
                .OfCategory(BuiltInCategory.OST_FabricationDuctwork)
                .WhereElementIsNotElementType()
                .ToElements())

    all_duct = list(fab_duct)
    if not all_duct:
        output.print_md("## No fabrication ducts found in view")
        script.exit()

    # Build parameter -> value -> elements map
    # Key is combination of all checked parameters
    param_groups = {}
    param_values_map = {}  # Store param values for later display

    for d in all_duct:
        param_values = {}
        skip_element = False

        # Get values for all parameters in check_parameters
        for param_name in check_parameters:
            param = lookup_parameter_case_insensitive(d, param_name)
            value = get_param_value(param) if param else None

            # Clean and normalize the value
            if value is not None:
                value_str = str(value).strip().lower()
            else:
                value_str = None

            param_values[param_name] = value_str

            # Skip if empty or in skip_values
            empty_vals = {None, "", "**"}
            if value_str in empty_vals:
                skip_element = True
                break
            if value_str in skip_values:
                skip_element = True
                break

        if skip_element:
            continue

        # Group by combination of all parameters (use cleaned values)
        composite_key = " | ".join(str(param_values[p]).strip(
        ) if param_values[p] else "" for p in sorted(check_parameters))
        if composite_key not in param_groups:
            param_groups[composite_key] = []
        param_groups[composite_key].append(d)
        param_values_map[d.Id.IntegerValue] = param_values

    if not param_groups:
        output.print_md(
            "## No ducts found with parameters populated: {}".format(", ".join(check_parameters)))
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

    # Sort by Item Number parameter value
    def get_sort_key(duct):
        values = param_values_map.get(duct.element.Id.IntegerValue, {})
        item_num = values.get("item number", "0")
        try:
            return int(item_num) if item_num else 0
        except ValueError:
            return 0

    fil_ducts.sort(key=get_sort_key)

    for i, fil in enumerate(fil_ducts, start=1):
        try:
            values = param_values_map.get(fil.element.Id.IntegerValue, {})
            param_str = " | ".join("{}: {}".format(p, values.get(p, "")) for p in sorted(check_parameters))
        except Exception:
            param_str = ""
        output.print_md(
            '### No: {:03} | ID: {} | {} | Length: {:06.2f}" | Size: {}'.format(
                i,
                output.linkify(fil.element.Id),
                param_str,
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
