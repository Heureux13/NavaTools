# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from Autodesk.Revit.DB import FilteredElementCollector, FabricationPart, IndependentTag, ElementId, View, ViewType, BuiltInCategory, ReferencePlane
from pyrevit import revit, forms
import sys

# Button info
# ===================================================
__title__ = "Isolate ReA"
__doc__ = """
Isolate Releaf Exhaust Air
"""

# Variables
# ==================================================
doc = revit.doc
active_view = doc.ActiveView
TARGET_SERVICE_CODE = "rea".strip().lower()


def get_service_code(service_name):
    """Return token between first and second dashes, stripped and lowercased."""
    if not service_name:
        return None
    parts = [p.strip() for p in service_name.split('-')]
    if len(parts) < 2 or not parts[1]:
        return None
    return parts[1].strip().lower()

# Main Code
# =================================================


# Collect all fabrication parts in the view
fab_collector = FilteredElementCollector(doc, active_view.Id)\
    .OfClass(FabricationPart)\
    .WhereElementIsNotElementType()

# Group elements by parsed service code (second token)
element_by_code = {}

for elem in fab_collector:
    try:
        service_name = elem.ServiceName
        service_code = get_service_code(service_name)
        if service_code:
            if service_code not in element_by_code:
                element_by_code[service_code] = []
            element_by_code[service_code].append(elem.Id)
    except BaseException:
        pass

if not element_by_code:
    forms.alert(
        "No fabrication services with dash-delimited codes were found in current view.",
        exitscript=True
    )

selected_element_ids = element_by_code.get(TARGET_SERVICE_CODE, [])
if not selected_element_ids:
    forms.alert(
        "No fabrication elements found with service code '{}' in current view.".format(
            TARGET_SERVICE_CODE
        ),
        exitscript=True
    )


def _id_int(eid):
    try:
        return eid.IntegerValue
    except Exception:
        try:
            return eid.Value
        except Exception:
            return None


def collect_related_tags(doc, view, host_ids):
    host_set = set()
    for eid in host_ids:
        v = _id_int(eid)
        if v is not None:
            host_set.add(v)

    tag_ids = []
    tags = (FilteredElementCollector(doc, view.Id)
            .OfClass(IndependentTag)
            .WhereElementIsNotElementType()
            .ToElements())

    for tag in tags:
        try:
            refs = []
            try:
                refs = list(tag.GetTaggedLocalElementIds() or [])
            except Exception:
                pass
            if not refs:
                teid = getattr(tag, 'TaggedElementId', None)
                if teid and teid != ElementId.InvalidElementId:
                    refs = [teid]
            for rid in refs:
                if _id_int(rid) in host_set:
                    tag_ids.append(tag.Id)
                    break
        except Exception:
            continue
    return tag_ids


# Collect element IDs for selected services
with revit.Transaction("Isolate by Fabrication Service"):
    element_ids = list(selected_element_ids)

    # Include related tags in isolation
    tag_ids = collect_related_tags(doc, active_view, element_ids)
    element_ids.extend(tag_ids)

    # Include annotation elements to keep them visible
    categories_to_include = [
        BuiltInCategory.OST_Grids,
        BuiltInCategory.OST_Dimensions,
        BuiltInCategory.OST_Viewers,
        BuiltInCategory.OST_MechanicalEquipment,
        BuiltInCategory.OST_MechanicalEquipmentTags,
    ]

    for bic in categories_to_include:
        collector = FilteredElementCollector(doc, active_view.Id).OfCategory(
            bic).WhereElementIsNotElementType()
        for elem in collector:
            element_ids.append(elem.Id)

    # Include reference planes
    ref_plane_collector = FilteredElementCollector(
        doc, active_view.Id).OfClass(ReferencePlane)
    for plane in ref_plane_collector:
        element_ids.append(plane.Id)

    # Apply temporary isolation to view
    if element_ids:
        from System.Collections.Generic import List
        id_list = List[ElementId](element_ids)
        active_view.IsolateElementsTemporary(id_list)
    else:
        print("No elements found for selected services")
