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
from Autodesk.Revit.DB import (BuiltInCategory, FilteredElementCollector, ElementId,
                               TemporaryViewMode, ReferencePlane)
from System.Collections.Generic import List

# Button info
# ===================================================
__title__ = "Workset Mechanical Equipment"
__doc__ = """
Toggle isolation and report worksets for MEP elements in the active view.
"""

# Variables
# ==================================================
doc = revit.doc
active_view = doc.ActiveView
output = script.get_output()

# Categories to isolate
categories_to_isolate = [
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_MechanicalEquipmentTags,
    BuiltInCategory.OST_Walls,
    BuiltInCategory.OST_StructuralFraming,
    BuiltInCategory.OST_Floors,
    BuiltInCategory.OST_Ceilings,
    BuiltInCategory.OST_Lines,
    BuiltInCategory.OST_Viewers,
    BuiltInCategory.OST_Dimensions,
    BuiltInCategory.OST_Grids,
    BuiltInCategory.OST_GenericAnnotation,
]

# Helpers
# ==================================================================================================


def collect_elements_from_categories(doc, view_id, categories):
    """Collect element IDs from specified categories in current document."""
    ids = List[ElementId]()

    excluded_types = ['SectionMarker', 'ElevationMarker', 'ViewSection']

    for bic in categories:
        collector = FilteredElementCollector(doc, view_id).OfCategory(
            bic).WhereElementIsNotElementType()
        for el in collector:
            element_type = el.GetType().Name
            if element_type not in excluded_types:
                ids.Add(el.Id)

    ref_plane_collector = FilteredElementCollector(doc, view_id).OfClass(ReferencePlane)
    for plane in ref_plane_collector:
        ids.Add(plane.Id)

    return ids


def is_view_isolated(view):
    """Check if view currently has isolation enabled."""
    try:
        return len(view.GetIsolatedElementIds()) > 0
    except BaseException:
        return False


def get_workset_name(doc, element):
    """Get the element workset name safely."""
    try:
        workset = doc.GetWorksetTable().GetWorkset(element.WorksetId)
        if workset:
            return workset.Name
    except BaseException:
        pass
    return 'Unknown Workset'


def print_isolated_elements_with_worksets(doc, element_ids):
    """Print isolated elements with linkified ids and workset names."""
    if element_ids.Count == 0:
        return

    grouped_by_workset = {}
    for elid in element_ids:
        element = doc.GetElement(elid)
        if not element:
            continue
        workset_name = get_workset_name(doc, element)
        grouped_by_workset.setdefault(workset_name, []).append(elid)

    output.print_md('### Isolated Elements Grouped by Workset')
    for workset_name in sorted(grouped_by_workset.keys()):
        grouped_ids = grouped_by_workset[workset_name]
        output.print_md('- {} | Count: {} | Select All: {}'.format(
            workset_name,
            len(grouped_ids),
            output.linkify(grouped_ids)
        ))

    output.print_md('### Isolated Elements and Worksets')
    for elid in element_ids:
        element = doc.GetElement(elid)
        if not element:
            continue

        workset_name = get_workset_name(doc, element)
        category_name = element.Category.Name if element.Category else 'No Category'
        output.print_md('- {} | {} | Workset: {}'.format(
            output.linkify(elid),
            category_name,
            workset_name
        ))


# Main Code
# =================================================
with revit.Transaction('Toggle Isolation'):
    if is_view_isolated(active_view):
        # Remove isolation
        active_view.DisableTemporaryViewMode(
            TemporaryViewMode.TemporaryIsolate)
    else:
        # Collect elements visible in current view only
        ids = collect_elements_from_categories(
            doc, active_view.Id, categories_to_isolate)

        # Apply isolation if we have elements
        if ids.Count > 0:
            active_view.IsolateElementsTemporary(ids)
            print_isolated_elements_with_worksets(doc, ids)
        else:
            # Show message if no elements found
            from pyrevit import forms
            forms.alert('No elements found to isolate.')
