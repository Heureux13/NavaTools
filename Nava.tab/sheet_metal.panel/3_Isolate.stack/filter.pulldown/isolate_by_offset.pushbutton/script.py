# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, ElementId, VisibleInViewFilter, ReferencePlane
from Autodesk.Revit.UI import TaskDialog
from pyrevit import revit, forms
from System.Collections.Generic import List

# Button info
# ===================================================
__title__ = "Isolate by Offset"
__doc__ = """
Isolate fabrication offset fittings in the active view.
"""

# Variables
# ==================================================
doc = revit.doc
active_view = revit.active_view

# Families targeted by the Offset Data script.
FAMILY_LIST = {
    'offset',
    'gored elbow',
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

element_ids = []
for d in all_duct:
    family_type = doc.GetElement(d.GetTypeId())
    if not family_type:
        continue
    family_name = family_type.FamilyName.lower().replace('*', '').strip()
    if family_name in FAMILY_LIST:
        element_ids.append(d.Id)

if not element_ids:
    forms.alert("No offset elements found in current view.", exitscript=True)

with revit.Transaction("Isolate Offsets"):
    # Include annotation elements to keep them visible
    categories_to_include = [
        BuiltInCategory.OST_Grids,
        BuiltInCategory.OST_Dimensions,
        BuiltInCategory.OST_Viewers,
        BuiltInCategory.OST_MechanicalEquipment,
        BuiltInCategory.OST_MechanicalEquipmentTags,
    ]

    for bic in categories_to_include:
        collector = FilteredElementCollector(doc, active_view.Id).OfCategory(bic).WhereElementIsNotElementType()
        for elem in collector:
            element_ids.append(elem.Id)

    # Include reference planes
    for plane in FilteredElementCollector(doc, active_view.Id).OfClass(ReferencePlane):
        element_ids.append(plane.Id)

    id_list = List[ElementId](element_ids)
    active_view.IsolateElementsTemporary(id_list)
