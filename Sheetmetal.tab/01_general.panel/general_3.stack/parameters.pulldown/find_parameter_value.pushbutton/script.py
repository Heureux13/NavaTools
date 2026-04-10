# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import revit, script
from config.parameters_registry import *
from Autodesk.Revit.DB import (
    BuiltInCategory,
    ElementId,
    FilteredElementCollector,
)
from System.Collections.Generic import List

# Button info
# ======================================================================
__title__ = 'Find Parameter Value'
__doc__ = '''
Find MEP elements with specified parameter and filter by value.
'''

# Configuration
# ======================================================================
parameter = '_type'
parameter_value = "sleeve"

# Code
# ======================================================================
doc = revit.doc
view = revit.active_view
uidoc = revit.uidoc
output = script.get_output()


def _element_id_value(eid):
    if eid is None:
        return None
    return eid.IntegerValue if hasattr(eid, "IntegerValue") else eid.Value


def _lookup_parameter_case_insensitive(element, param_name):
    param_name_lower = param_name.strip().lower()

    for param in element.Parameters:
        try:
            if param.Definition.Name.strip().lower() == param_name_lower:
                return param
        except Exception:
            pass

    try:
        element_type = doc.GetElement(element.GetTypeId())
    except Exception:
        element_type = None

    if element_type:
        for param in element_type.Parameters:
            try:
                if param.Definition.Name.strip().lower() == param_name_lower:
                    return param
            except Exception:
                pass

    return None


def _get_parameter_text(param):
    if param is None:
        return ""

    try:
        if param.StorageType == 0:
            return ""
    except Exception:
        pass

    try:
        value = param.AsString()
        if value:
            return value.strip()
    except Exception:
        pass

    try:
        value = param.AsValueString()
        if value:
            return value.strip()
    except Exception:
        pass

    try:
        if param.StorageType == 1:
            return str(param.AsDouble()).strip()
        if param.StorageType == 2:
            return str(param.AsInteger()).strip()
        if param.StorageType == 3:
            return str(_element_id_value(param.AsElementId())).strip()
    except Exception:
        pass

    return ""


# Collect all MEP elements
categories = [
    BuiltInCategory.OST_FabricationDuctwork,
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_DuctTerminal,
]

all_elements = []
for cat in categories:
    all_elements.extend(
        FilteredElementCollector(doc, view.Id)
        .OfCategory(cat)
        .WhereElementIsNotElementType()
        .ToElements()
    )

if not all_elements:
    output.print_md("## No MEP elements found in this view.")
    script.exit()

output.print_md("## MEP Elements with '{}'".format(parameter))
output.print_md("Total elements: {}".format(len(all_elements)))

# Find elements with the parameter
all_with_param = {}
matching_elements = {}
for elem in all_elements:
    param = _lookup_parameter_case_insensitive(elem, parameter)
    if param:
        value = _get_parameter_text(param)

        all_with_param[elem.Id] = (elem, value)
        # Filter by the specified value
        if value.lower() == parameter_value.lower():
            matching_elements[elem.Id] = (elem, value)

output.print_md("### Elements with '{}' parameter: {}".format(parameter, len(all_with_param)))

if matching_elements:
    grouped_elements = {}
    selected_elements = []

    for elem_id, (elem, value) in matching_elements.items():
        grouped_elements.setdefault(value, []).append(elem)
        selected_elements.append(elem)

    id_list = List[ElementId]([e.Id for e in selected_elements])
    uidoc.Selection.SetElementIds(id_list)

    index = 1
    for grouped_value in sorted(grouped_elements.keys(), key=lambda x: x.lower()):
        output.print_md("## {} ({})".format(grouped_value, len(grouped_elements[grouped_value])))
        output.print_md("---")

        for elem in grouped_elements[grouped_value]:
            family_type = doc.GetElement(elem.GetTypeId())
            family_name = family_type.FamilyName if family_type else "Unknown"

            output.print_md(
                "### Index: {:03} | Element ID: {} | Family: {} | {}: {}".format(
                    index,
                    output.linkify(elem.Id),
                    family_name,
                    parameter,
                    grouped_value,
                )
            )
            index += 1

    output.print_md("---")
    output.print_md(
        "### Selected Element IDs: {}".format(
            output.linkify([e.Id for e in selected_elements])
        )
    )
    output.print_md("---")
    output.print_md("# Total Elements Selected: {}".format(len(selected_elements)))
else:
    output.print_md("\n### No exact match for '{}'. All values:".format(parameter_value))
    for elem_id, (elem, value) in sorted(all_with_param.items()):
        output.print_md("- {} | Value: '{}' | {}".format(
            output.linkify(elem_id),
            value,
            getattr(elem, 'Name', 'Unknown')
        ))
