# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from Autodesk.Revit.DB import FilteredElementCollector, FabricationPart, IndependentTag, ElementId
from pyrevit import revit, forms
import sys

# Button info
# ===================================================
__title__ = "Isolate by Service"
__doc__ = """
Isolates view to show only fabrication elements with selected services
"""

# Variables
# ==================================================
doc = revit.doc
active_view = doc.ActiveView

# Main Code
# =================================================

# Collect all fabrication parts in the view
fab_collector = FilteredElementCollector(doc, active_view.Id)\
    .OfClass(FabricationPart)\
    .WhereElementIsNotElementType()

# Get unique services from all fabrication parts
services = set()
element_by_service = {}

for elem in fab_collector:
    try:
        service_name = elem.ServiceName
        if service_name:
            services.add(service_name)
            if service_name not in element_by_service:
                element_by_service[service_name] = []
            element_by_service[service_name].append(elem.Id)
    except BaseException:
        pass

if not services:
    forms.alert("No fabrication services found in current view.",
                exitscript=True)

# Show selection dialog
selected_services = forms.SelectFromList.show(
    sorted(services),
    title='Select Fabrication Services to Isolate',
    multiselect=True,
    button_name='Isolate Selected'
)

if not selected_services:
    sys.exit(0)


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
    element_ids = []
    for service in selected_services:
        element_ids.extend(element_by_service[service])

    # Include related tags in isolation
    tag_ids = collect_related_tags(doc, active_view, element_ids)
    element_ids.extend(tag_ids)

    # Apply temporary isolation to view
    if element_ids:
        from System.Collections.Generic import List
        id_list = List[ElementId](element_ids)
        active_view.IsolateElementsTemporary(id_list)
    else:
        print("No elements found for selected services")
