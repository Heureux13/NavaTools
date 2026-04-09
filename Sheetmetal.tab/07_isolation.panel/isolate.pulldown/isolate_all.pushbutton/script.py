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
from config.parameters_registry import RVT_CLEARANCE_ZONE, RVT_FAMILY, RVT_TYPE

# Button info
# ===================================================
__title__ = "Isolate by MEP"
__doc__ = """
Toggle isolation of walls, ducts, pipes, steel beams, and floors.
"""

# Variables
# ==================================================
doc = revit.doc
active_view = doc.ActiveView
output = script.get_output()

# Categories to isolate
categories_to_isolate = [
    BuiltInCategory.OST_Ceilings,
    BuiltInCategory.OST_Dimensions,
    BuiltInCategory.OST_DuctAccessory,
    BuiltInCategory.OST_DuctCurves,
    BuiltInCategory.OST_DuctFitting,
    BuiltInCategory.OST_DuctInsulations,
    BuiltInCategory.OST_DuctTags,
    BuiltInCategory.OST_DuctTerminal,
    BuiltInCategory.OST_DuctTerminalTags,
    BuiltInCategory.OST_FabricationContainment,
    BuiltInCategory.OST_FabricationContainmentCenterLine,
    BuiltInCategory.OST_FabricationContainmentDrop,
    BuiltInCategory.OST_FabricationContainmentRise,
    BuiltInCategory.OST_FabricationContainmentSymbology,
    BuiltInCategory.OST_FabricationContainmentTags,
    BuiltInCategory.OST_FabricationDuctwork,
    BuiltInCategory.OST_FabricationDuctworkTags,
    BuiltInCategory.OST_FabricationHangerTags,
    BuiltInCategory.OST_FabricationHangers,
    BuiltInCategory.OST_FabricationPipework,
    BuiltInCategory.OST_FabricationPipeworkCenterLine,
    BuiltInCategory.OST_FabricationPipeworkDrop,
    BuiltInCategory.OST_FabricationPipeworkInsulation,
    BuiltInCategory.OST_FabricationPipeworkRise,
    BuiltInCategory.OST_FabricationPipeworkSymbology,
    BuiltInCategory.OST_FabricationPipeworkTags,
    BuiltInCategory.OST_FabricationServiceElements,
    BuiltInCategory.OST_FlexDuctCurves,
    BuiltInCategory.OST_FlexDuctTags,
    BuiltInCategory.OST_FlexPipeCurves,
    BuiltInCategory.OST_FlexPipeTags,
    BuiltInCategory.OST_Floors,
    BuiltInCategory.OST_GenericAnnotation,
    BuiltInCategory.OST_GenericModel,
    BuiltInCategory.OST_Grids,
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_MechanicalEquipmentTags,
    BuiltInCategory.OST_PipeAccessory,
    BuiltInCategory.OST_PipeCurves,
    BuiltInCategory.OST_PipeFitting,
    BuiltInCategory.OST_PipeInsulations,
    BuiltInCategory.OST_PipeTags,
    BuiltInCategory.OST_StructuralFraming,
    BuiltInCategory.OST_Viewers,
    BuiltInCategory.OST_Walls,
]

# Helpers
# ==================================================================================================


def collect_elements_from_categories(doc, view_id, categories):
    """Collect element IDs from specified categories in current document."""
    ids = List[ElementId]()

    # Element types to keep visible (not isolate)
    excluded_types = ['SectionMarker', 'ElevationMarker', 'ViewSection']

    for bic in categories:
        collector = FilteredElementCollector(doc, view_id).OfCategory(
            bic).WhereElementIsNotElementType()
        for el in collector:
            # Skip annotation elements - keep them visible
            element_type = el.GetType().Name
            if element_type not in excluded_types:
                ids.Add(el.Id)

    # Collect reference planes separately (no BuiltInCategory)
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


def get_element_type(element):
    """Get element type safely."""
    try:
        type_id = element.GetTypeId()
        if type_id and type_id != ElementId.InvalidElementId:
            return doc.GetElement(type_id)
    except BaseException:
        pass
    return None


def get_text(value):
    """Normalize text values for matching."""
    if value is None:
        return ""
    return str(value).strip().lower()


def get_parameter_text(element, name):
    """Return normalized parameter text value or empty string."""
    try:
        param = element.LookupParameter(name)
        if not param:
            return ""
        val = param.AsString() or param.AsValueString()
        return get_text(val)
    except BaseException:
        return ""


def is_clearance_like_element(element):
    """Match clearance elements by category plus family/type/zone markers."""
    if not element or not element.Category:
        return False

    cat_id = element.Category.Id.IntegerValue
    allowed_categories = [
        int(BuiltInCategory.OST_MechanicalEquipment),
        int(BuiltInCategory.OST_GenericModel),
    ]
    if cat_id not in allowed_categories:
        return False

    element_type = get_element_type(element)
    family_name = get_parameter_text(element, RVT_FAMILY)
    type_name = get_parameter_text(element, RVT_TYPE)

    if not family_name:
        try:
            if element_type and getattr(element_type, 'FamilyName', None):
                family_name = get_text(element_type.FamilyName)
        except BaseException:
            pass

    if not type_name:
        try:
            if element_type and getattr(element_type, 'Name', None):
                type_name = get_text(element_type.Name)
        except BaseException:
            pass

    if 'clearance' in family_name or 'clearance' in type_name:
        return True

    clearance_zone = get_parameter_text(element, RVT_CLEARANCE_ZONE)
    return clearance_zone in ['yes', '1', 'true']


def collect_clearance_like_elements(doc, view_id):
    """Collect matching clearance elements in the active view."""
    ids = List[ElementId]()
    collector = FilteredElementCollector(doc, view_id).WhereElementIsNotElementType()

    for el in collector:
        if is_clearance_like_element(el):
            ids.Add(el.Id)

    return ids


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

        # Ensure clearance-like instances are included.
        clearance_ids = collect_clearance_like_elements(doc, active_view.Id)
        for clearance_id in clearance_ids:
            ids.Add(clearance_id)

        # Apply isolation if we have elements
        if ids.Count > 0:
            active_view.IsolateElementsTemporary(ids)
        else:
            # Show message if no elements found
            output.print_md('No elements found to isolate.')
