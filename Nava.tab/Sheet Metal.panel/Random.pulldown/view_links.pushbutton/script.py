# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from pyrevit import revit, forms
from Autodesk.Revit.DB import FilteredElementCollector, RevitLinkInstance, ElementId
from System.Collections.Generic import List
import sys

# Button info
# ===================================================
__title__ = "Hide Links"
__doc__ = """
Hide selected Revit links in the active view (temporary hide).
"""

# Variables
# ==================================================
doc = revit.doc
active_view = doc.ActiveView

links_in_view = list(
    FilteredElementCollector(doc, active_view.Id)
    .OfClass(RevitLinkInstance)
    .WhereElementIsNotElementType()
)

if not links_in_view:
    forms.alert('No Revit links found in this view.')
    sys.exit(0)

display_map = {}
for link in links_in_view:
    display_name = '{} (Id: {})'.format(link.Name, link.Id.IntegerValue)
    display_map[display_name] = link.Id

selected_names = forms.SelectFromList.show(
    sorted(display_map.keys()),
    title='Select Links to Hide',
    multiselect=True,
    button_name='Hide Selected'
)

if not selected_names:
    sys.exit(0)

ids = List[ElementId]()
for name in selected_names:
    link_id = display_map.get(name)
    if link_id:
        ids.Add(link_id)

if ids.Count == 0:
    sys.exit(0)

with revit.Transaction('Hide Selected Links'):
    active_view.HideElementsTemporary(ids)
