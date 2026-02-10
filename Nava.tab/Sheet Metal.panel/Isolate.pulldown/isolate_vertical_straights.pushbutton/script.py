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
from revit_xyz import RevitXYZ

# Button info
# ===================================================
__title__ = "Isolate Vertical Straights"
__doc__ = """
Toggle isolation to show only vertical straight duct elements.
"""

# Variables
# ==================================================
doc = revit.doc
active_view = revit.active_view

# Angle threshold for vertical ducts (degrees)
VERTICAL_THRESHOLD = 80.0

# Helpers
# ==================================================================================================
straigth_group = {
    "Straight",
}


def collect_straight_ducts(doc, view):
    """Collect duct element IDs where family equals 'Straight' and angle is vertical."""
    ids = List[ElementId]()

    # Get all fabrication ducts in the view
    ducts = RevitDuct.all(doc, view)

    # Filter to only vertical straights
    for duct in ducts:
        if duct.family in straigth_group:
            # Check if duct is vertical using RevitXYZ
            xyz = RevitXYZ(duct.element)
            angle = xyz.straight_joint_degree()

            # Only add if angle is close to vertical (±90 degrees)
            if angle is not None and abs(angle) >= VERTICAL_THRESHOLD:
                ids.Add(duct.element.Id)

    return ids


def is_view_isolated(view):
    """Check if view currently has isolation enabled."""
    try:
        return len(view.GetIsolatedElementIds()) > 0
    except BaseException:
        return False


# Main Code
# =================================================
with revit.Transaction('Toggle Isolate Straights'):
    if is_view_isolated(active_view):
        # Remove isolation
        active_view.DisableTemporaryViewMode(
            TemporaryViewMode.TemporaryIsolate)
    else:
        # Collect straight ducts visible in current view only
        ids = collect_straight_ducts(doc, active_view)

        # Isolate straight ducts if we have elements
        if ids.Count > 0:
            active_view.IsolateElementsTemporary(ids)
