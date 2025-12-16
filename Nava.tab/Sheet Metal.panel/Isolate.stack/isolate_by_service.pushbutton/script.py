# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from Autodesk.Revit.DB import FilteredElementCollector, FabricationPart
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
    except:
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

# Collect element IDs for selected services
with revit.Transaction("Isolate by Fabrication Service"):
    element_ids = []
    for service in selected_services:
        element_ids.extend(element_by_service[service])

    # Apply temporary isolation to view
    if element_ids:
        from System.Collections.Generic import List
        from Autodesk.Revit.DB import ElementId
        id_list = List[ElementId](element_ids)
        active_view.IsolateElementsTemporary(id_list)
        print("Isolated {} elements from {} services".format(
            len(element_ids), len(selected_services)))
    else:
        print("No elements found for selected services")
