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

matched_elements = []
for d in all_duct:
    family_type = doc.GetElement(d.GetTypeId())
    if not family_type:
        continue
    family_name = family_type.FamilyName.lower().replace('*', '').strip()
    if family_name in FAMILY_LIST:
        matched_elements.append(d)

if not matched_elements:
    forms.alert("No offset elements found in current view.", exitscript=True)

# Select all matched offsets
id_list = List[ElementId]([e.Id for e in matched_elements])
uidoc.Selection.SetElementIds(id_list)

# Print output
for i, elem in enumerate(matched_elements, start=1):
    family_type = doc.GetElement(elem.GetTypeId())
    family_name = family_type.FamilyName if family_type else "Unknown"
    offset_param = elem.LookupParameter(PYT_OFFSET_VALUE)
    offset_str = "-"
    if offset_param:
        offset_str = offset_param.AsString() or offset_param.AsValueString() or "-"

    output.print_md(
        '### Index: {:03} | Element ID: {} | Offset: {} | Family: {}'.format(
            i,
            output.linkify(elem.Id),
            offset_str,
            family_name,
        )
    )

output.print_md("---")
output.print_md("# Total Elements Selected: {}".format(len(matched_elements)))
