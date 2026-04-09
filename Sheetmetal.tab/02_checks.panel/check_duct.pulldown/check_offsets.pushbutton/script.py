# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, ElementId, VisibleInViewFilter
from pyrevit import revit, forms, script
from System.Collections.Generic import List
from config.parameters_registry import *

# Button info
# ===================================================
__title__ = "Offsets"
__doc__ = """
Select fabrication offset fittings in the active view and print a summary.
"""

# Variables
# ==================================================
doc = revit.doc
active_view = revit.active_view
uidoc = __revit__.ActiveUIDocument
output = script.get_output()

# Families targeted by the Offset Data script.
FAMILY_LIST = {
    'offset',
    'ogee',
    'oval reducer',
    'oval to round',
    'reducer',
    'square to ø',
    'transition',
    'cid330 - (radius 2-way offset)',
}


# Main Code
# ==================================================


all_duct = (FilteredElementCollector(doc, active_view.Id)
            .OfCategory(BuiltInCategory.OST_FabricationDuctwork)
            .WhereElementIsNotElementType()
            .WherePasses(VisibleInViewFilter(doc, active_view.Id))
            .ToElements())

grouped_elements = {}
for d in all_duct:
    family_type = doc.GetElement(d.GetTypeId())
    if not family_type:
        continue
    family_name = family_type.FamilyName.lower().replace('*', '').strip()
    if family_name not in FAMILY_LIST:
        continue

    offset_param = d.LookupParameter(PYT_OFFSET_VALUE) or d.LookupParameter(LEGACY_OFFSET)
    offset_str = ""
    if offset_param:
        offset_str = offset_param.AsString() or offset_param.AsValueString() or ""
    offset_str = str(offset_str).strip()

    if offset_str.upper() == "CL":
        continue

    group_key = offset_str if offset_str else "(blank)"
    if group_key not in grouped_elements:
        grouped_elements[group_key] = []
    grouped_elements[group_key].append(d)

if not grouped_elements:
    forms.alert("No non-CL offset elements found in current view.", exitscript=True)

# Select all grouped offsets
selected_elements = []
for elements in grouped_elements.values():
    selected_elements.extend(elements)

id_list = List[ElementId]([e.Id for e in selected_elements])
uidoc.Selection.SetElementIds(id_list)

# Print output
index = 1
for offset_value in sorted(grouped_elements.keys(), key=lambda x: x.lower()):
    output.print_md("## {} ({})".format(offset_value, len(grouped_elements[offset_value])))
    output.print_md("---")

    for elem in grouped_elements[offset_value]:
        family_type = doc.GetElement(elem.GetTypeId())
        family_name = family_type.FamilyName if family_type else "Unknown"

        output.print_md(
            '### Index: {:03} | Element ID: {} | Family: {} | Offset: {}'.format(
                index,
                output.linkify(elem.Id),
                family_name,
                offset_value,
            )
        )
        index += 1

output.print_md("---")
output.print_md("# Total Elements Selected: {}".format(len(selected_elements)))
