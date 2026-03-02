# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from pyrevit import revit, script
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    ElementId,
    TemporaryViewMode,
)
from System.Collections.Generic import List
import sys

# Button info
# ===================================================
__title__ = "Hide Hangers"
__doc__ = """
Hides all fabrication hangers in active view.
"""

# Variables
# ==================================================
doc = revit.doc
active_view = doc.ActiveView
output = script.get_output()

# Collect all fabrication hangers in the document
hangers = list(
    FilteredElementCollector(doc, active_view.Id)
    .OfCategory(BuiltInCategory.OST_FabricationHangers)
    .WhereElementIsNotElementType()
)

if not hangers:
    output.print_md('**No fabrication hangers found in this model.**')
    sys.exit(0)

# Build list of hanger IDs to hide
ids_to_hide = List[ElementId]()

# Add hanger IDs
for hanger in hangers:
    ids_to_hide.Add(hanger.Id)

# Apply temporary hide (without clearing existing hidden elements)
with revit.Transaction('Hide All Hangers'):
    try:
        if ids_to_hide.Count > 0:
            active_view.HideElementsTemporary(ids_to_hide)
            output.print_md('**Hidden {} fabrication hangers.**'.format(len(hangers)))
    except Exception as e:
        output.print_md('**This view does not allow temporary hiding: {}**'.format(str(e)))
