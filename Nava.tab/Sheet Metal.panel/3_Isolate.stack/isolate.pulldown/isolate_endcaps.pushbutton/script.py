# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from pyrevit import revit
from Autodesk.Revit.DB import TemporaryViewMode, ElementId, BuiltInCategory, FilteredElementCollector
from System.Collections.Generic import List
from revit_duct import RevitDuct

# Button info
# ===================================================
__title__ = "Isolate Elbows"
__doc__ = """
Toggle isolation to show only elbow duct elements.
"""

# Variables
# ==================================================
doc = revit.doc
active_view = revit.active_view

# Helpers
# ==================================================================================================

group = {"Elbow",
         "Elbow - 90 Degree",
         }


def collect_elbow_ducts(doc, view):
    """Collect duct element IDs where family is Elbow plus all hangers."""
    ids = List[ElementId]()

    # Get all fabrication ducts in the view
    ducts = RevitDuct.all(doc, view)

    # Filter to only elbows
    for duct in ducts:
        if duct.family in group:
            ids.Add(duct.element.Id)

    # Also collect all hangers
    hangers = FilteredElementCollector(doc, view.Id).OfCategory(
        BuiltInCategory.OST_FabricationHangers).WhereElementIsNotElementType().ToElements()
    for hanger in hangers:
        ids.Add(hanger.Id)

    return ids


def is_view_isolated(view):
    """Check if view currently has isolation enabled."""
    try:
        return len(view.GetIsolatedElementIds()) > 0
    except BaseException:
        return False


# Main Code
# =================================================
with revit.Transaction('Toggle Isolate Elbows'):
    if is_view_isolated(active_view):
        # Remove isolation
        active_view.DisableTemporaryViewMode(
            TemporaryViewMode.TemporaryIsolate)
    else:
        # Collect elbow ducts visible in current view only
        ids = collect_elbow_ducts(doc, active_view)

        # Isolate elbow ducts if we have elements
        if ids.Count > 0:
            active_view.IsolateElementsTemporary(ids)
