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
from Autodesk.Revit.DB import FilteredElementCollector, ElementId, TemporaryViewMode
from System.Collections.Generic import List

# Button info
# ===================================================
__title__ = "Isolate by Matchlines"
__doc__ = """
Toggle isolation of matchlines and view references in the current view.
"""

# Variables
# ==================================================
doc = revit.doc
active_view = doc.ActiveView

# Helpers
# ==================================================================================================


def collect_matchlines(doc, view_id):
    """Collect matchline and view reference element IDs from the current view."""
    ids = List[ElementId]()

    # Get the Matchline and View References categories by name
    category_names = ["Matchline", "View References"]

    for cat_name in category_names:
        try:
            category = doc.Settings.Categories.get_Item(cat_name)
            if category:
                collector = FilteredElementCollector(doc, view_id).OfCategoryId(category.Id)
                for el in collector:
                    ids.Add(el.Id)
        except BaseException:
            pass

    return ids


def is_view_isolated(view):
    """Check if view currently has isolation enabled."""
    try:
        return len(view.GetIsolatedElementIds()) > 0
    except BaseException:
        return False


# Main Code
# =================================================
with revit.Transaction('Toggle Isolation'):
    if is_view_isolated(active_view):
        # Remove isolation
        active_view.DisableTemporaryViewMode(
            TemporaryViewMode.TemporaryIsolate)
    else:
        # Collect matchlines from current view only
        ids = collect_matchlines(doc, active_view.Id)

        # Apply isolation if we have elements
        if ids.Count > 0:
            active_view.IsolateElementsTemporary(ids)
        else:
            # Show message if no elements found
            from pyrevit import forms
            forms.alert('No matchlines or view references found to isolate.')
