# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from collections import defaultdict
from pyrevit import revit, forms
from Autodesk.Revit.DB import FilteredElementCollector, ElementId, View
from System.Collections.Generic import List
import sys

# Button info
# ===================================================
__title__ = "Select View References"
__doc__ = """
Shows all view references in the current view grouped by their target floor.
Select the groups you want to change, and they'll be selected in Revit so you
can batch-edit them in the Properties panel.
"""

# Variables
# ==================================================
doc = revit.doc
uidoc = __revit__.ActiveUIDocument
active_view = doc.ActiveView


# Helpers
# ==================================================================================================


def collect_view_references(doc, view_id):
    """Collect all View Reference elements in the given view, trying multiple strategies."""
    # Strategy 1: ViewReference class (Revit 2022+)
    try:
        from Autodesk.Revit.DB import ViewReference as VR
        elems = list(FilteredElementCollector(doc, view_id).OfClass(VR))
        if elems:
            return elems
    except Exception:
        pass

    # Strategy 2: BuiltInCategory OST_ReferenceViewer
    try:
        from Autodesk.Revit.DB import BuiltInCategory
        elems = list(
            FilteredElementCollector(doc, view_id)
            .OfCategory(BuiltInCategory.OST_ReferenceViewer)
            .WhereElementIsNotElementType()
        )
        if elems:
            return elems
    except Exception:
        pass

    # Strategy 3: Search all category names containing "ref" or "view"
    try:
        for cat in doc.Settings.Categories:
            name_lower = cat.Name.lower()
            if "view ref" in name_lower or "reference view" in name_lower:
                elems = list(
                    FilteredElementCollector(doc, view_id).OfCategoryId(cat.Id)
                )
                if elems:
                    return elems
    except Exception:
        pass

    # Strategy 4: Broad search — any element in view whose category has "ref" in it
    try:
        found = []
        for el in FilteredElementCollector(doc, view_id).WhereElementIsNotElementType():
            try:
                cat = el.Category
                if cat and "ref" in cat.Name.lower():
                    found.append(el)
            except Exception:
                pass
        if found:
            return found
    except Exception:
        pass

    return []


def get_target_view_id(elem):
    """Get the target view ElementId of a view reference element."""
    # Try ViewReference class method (Revit 2022+)
    try:
        from Autodesk.Revit.DB import ViewReference
        if isinstance(elem, ViewReference):
            vid = elem.GetReferencedViewId()
            if vid and vid != ElementId.InvalidElementId:
                return vid
    except Exception:
        pass

    # Parameter is named "Target view" (lowercase v) in this Revit version
    p = elem.LookupParameter("Target view")
    if p:
        try:
            eid = p.AsElementId()
            if eid and eid != ElementId.InvalidElementId:
                return eid
        except Exception:
            pass

    return None


def get_all_views(doc):
    """Return {view_name: view} for all non-template views in the document."""
    views = {}
    for v in FilteredElementCollector(doc).OfClass(View):
        try:
            if not v.IsTemplate:
                views[v.Name] = v
        except Exception:
            pass
    return views


# Main Code
# =================================================

refs = collect_view_references(doc, active_view.Id)
if not refs:
    forms.alert("No view references found in the current view.", exitscript=True)

all_views = get_all_views(doc)
view_by_id = {v.Id: v for v in all_views.values()}

# Group elements by current target view name
by_target = defaultdict(list)
for el in refs:
    target_id = get_target_view_id(el)
    if target_id and target_id in view_by_id:
        target_name = view_by_id[target_id].Name
        by_target[target_name].append(el.Id)

if not by_target:
    forms.alert(
        "Could not read the target view from any view reference.",
        exitscript=True
    )

# Build display list: "Floor Plan: 1st Floor... (x12)"
display_items = sorted(
    ["{:<50}  x{}".format(name, len(ids)) for name, ids in by_target.items()]
)

# Reverse map
name_from_display = {}
for name in by_target.keys():
    display = "{:<50}  x{}".format(name, len(by_target[name]))
    name_from_display[display] = name

# Show multi-select dialog
selected_displays = forms.SelectFromList.show(
    display_items,
    title="View References — Select to Change",
    multiselect=True,
    button_name="Select These"
)

if not selected_displays:
    sys.exit(0)

# Build final selection list
ids_to_select = List[ElementId]()
for display in selected_displays:
    name = name_from_display[display]
    for eid in by_target[name]:
        ids_to_select.Add(eid)

# Select them in Revit
uidoc.Selection.SetElementIds(ids_to_select)
