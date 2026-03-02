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
)
from System.Collections.Generic import List
import sys

# Button info
# ===================================================
__title__ = "Hide GRDs"
__doc__ = """
Hides all GRDs (air terminals) in active view.
"""

# Variables
# ==================================================
doc = revit.doc
active_view = doc.ActiveView
output = script.get_output()

grds = list(
    FilteredElementCollector(doc, active_view.Id)
    .OfCategory(BuiltInCategory.OST_DuctTerminal)
    .WhereElementIsNotElementType()
)

if not grds:
    output.print_md('**No GRDs found in this view.**')
    sys.exit(0)

# Build list of GRD IDs to hide
ids_to_hide = List[ElementId]()

# Add GRD IDs
for grd in grds:
    ids_to_hide.Add(grd.Id)

# Apply temporary hide (without clearing existing hidden elements)
with revit.Transaction('Hide GRDs'):
    try:
        if ids_to_hide.Count > 0:
            active_view.HideElementsTemporary(ids_to_hide)
            output.print_md('**Hidden {} GRDs.**'.format(len(grds)))
    except Exception:
        # output.print_md('**This view does not allow temporary hiding: {}**'.format(str(e)))
        pass
