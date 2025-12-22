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
from Autodesk.Revit.DB import (BuiltInCategory, FilteredElementCollector, ElementId,
                               TemporaryViewMode)
from System.Collections.Generic import List

# Button info
# ===================================================
__title__ = "Isolate All"
__doc__ = """
Toggle isolation of walls, ducts, pipes, steel beams, and floors.
Includes linked model components.
"""

# Variables
# ==================================================
doc = revit.doc
active_view = doc.ActiveView

# Categories to isolate
categories_to_isolate = [
    BuiltInCategory.OST_Walls,
    BuiltInCategory.OST_DuctCurves,
    BuiltInCategory.OST_FlexDuctCurves,
    BuiltInCategory.OST_DuctFitting,
    BuiltInCategory.OST_DuctAccessory,
    BuiltInCategory.OST_DuctTerminal,
    BuiltInCategory.OST_DuctInsulations,
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_FabricationDuctwork,
    BuiltInCategory.OST_FabricationHangers,
    BuiltInCategory.OST_PipeCurves,
    BuiltInCategory.OST_FlexPipeCurves,
    BuiltInCategory.OST_PipeFitting,
    BuiltInCategory.OST_PipeAccessory,
    BuiltInCategory.OST_PipeInsulations,
    BuiltInCategory.OST_StructuralFraming,
    BuiltInCategory.OST_Floors,
    BuiltInCategory.OST_Ceilings,
    BuiltInCategory.OST_Viewers,
]

# Main Code
# =================================================


def collect_elements_from_categories(doc, view_id, categories):
    """Collect element IDs from specified categories in current document."""
    ids = List[ElementId]()

    # Element types to keep visible (not isolate)
    excluded_types = ['SectionMarker', 'ElevationMarker', 'ViewSection', 'Grid', 'ReferencePlane']

    for bic in categories:
        collector = FilteredElementCollector(doc, view_id).OfCategory(
            bic).WhereElementIsNotElementType()
        for el in collector:
            # Skip annotation elements - keep them visible
            element_type = el.GetType().Name
            if element_type not in excluded_types:
                ids.Add(el.Id)
    return ids


def collect_linked_elements(doc, view_id, categories):
    """Collect linked instances that contain target categories."""
    ids = List[ElementId]()
    from Autodesk.Revit.DB import RevitLinkInstance

    # Get all linked instance elements
    linked_instances = FilteredElementCollector(
        doc, view_id).OfCategory(BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType()

    for link_inst in linked_instances:
        try:
            if isinstance(link_inst, RevitLinkInstance):
                linked_doc = link_inst.GetLinkDocument()
                if linked_doc:
                    # Check if linked doc has any of our target categories
                    has_target = False
                    for bic in categories:
                        collector = FilteredElementCollector(
                            linked_doc).OfCategory(bic).WhereElementIsNotElementType()
                        if collector.GetElementCount() > 0:
                            has_target = True
                            break
                    # If it has target elements, add the linked instance
                    if has_target:
                        ids.Add(link_inst.Id)
        except BaseException:
            pass
    return ids


def is_view_isolated(view):
    """Check if view currently has isolation enabled."""
    try:
        return len(view.GetIsolatedElementIds()) > 0
    except BaseException:
        return False


# Check if isolation is already active
with revit.Transaction('Toggle Isolation'):
    if is_view_isolated(active_view):
        # Remove isolation
        active_view.DisableTemporaryViewMode(TemporaryViewMode.TemporaryIsolate)
    else:
        # Collect elements from current document only
        ids = collect_elements_from_categories(doc, active_view.Id, categories_to_isolate)

        # Apply isolation if we have elements
        if ids.Count > 0:
            active_view.IsolateElementsTemporary(ids)
        else:
            # Show message if no elements found
            from pyrevit import forms
            forms.alert('No elements found to isolate.')
