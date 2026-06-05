# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.DB import FilteredElementCollector, ElementId, View, Viewport
from Autodesk.Revit.UI import RevitCommandId, PostableCommand
from System.Collections.Generic import List
from pyrevit import revit, script
import sys

# Button info
# ======================================================================
__title__ = 'Hide Section Views'
__doc__ = '''
Permanently hides section view markers/elements in selected views.
'''

# Variables
# ======================================================================
doc = revit.doc
uidoc = revit.uidoc
active_view = doc.ActiveView


def _get_element_id_value(element_id):
    try:
        return element_id.Value
    except Exception:
        pass

    try:
        return element_id.IntegerValue
    except Exception:
        pass

    try:
        return int(element_id)
    except Exception:
        return None


def _safe_param_text(element, param_name):
    try:
        param = element.LookupParameter(param_name)
        if not param:
            return ''
        value = param.AsString() or param.AsValueString()
        return (value or '').strip().lower()
    except Exception:
        return ''


def _is_section_view_element(element):
    try:
        category = element.Category
        if not category or category.Name != 'Views':
            return False
    except Exception:
        return False

    tokens = []

    try:
        tokens.append((getattr(element, 'Name', '') or '').lower())
    except Exception:
        pass

    for pname in ['Family', 'Family and Type', 'Type', 'View Name']:
        tokens.append(_safe_param_text(element, pname))

    combined = ' '.join([t for t in tokens if t])
    return 'section' in combined


def _is_own_section_reference(element, target_view):
    # Keep the marker/reference that points to the same view being processed.
    target_name = (getattr(target_view, 'Name', '') or '').strip().lower()
    if not target_name:
        return False

    for pname in ['View Name', 'Type', 'Family and Type', 'Name']:
        try:
            value = _safe_param_text(element, pname)
            if value and target_name in value:
                return True
        except Exception:
            pass

    try:
        element_name = (getattr(element, 'Name', '') or '').strip().lower()
        if element_name and target_name in element_name:
            return True
    except Exception:
        pass

    return False


def _zoom_extents():
    try:
        cmd_id = RevitCommandId.LookupPostableCommandId(PostableCommand.ZoomToFit)
        uidoc.Application.PostCommand(cmd_id)
    except Exception:
        pass


def _is_valid_target_view(view):
    if not view:
        return False

    try:
        if view.IsTemplate:
            return False
    except Exception:
        return False

    try:
        if str(view.ViewType) == 'DrawingSheet':
            return False
    except Exception:
        pass

    return True


def _get_target_views_from_selection_or_active():
    selected_ids = list(uidoc.Selection.GetElementIds())
    selected_views = {}

    for element_id in selected_ids:
        elem = doc.GetElement(element_id)
        if not elem:
            continue

        try:
            if isinstance(elem, Viewport):
                placed_view = doc.GetElement(elem.ViewId)
                if _is_valid_target_view(placed_view):
                    selected_views[_get_element_id_value(placed_view.Id)] = placed_view
                continue
        except Exception:
            pass

        try:
            if isinstance(elem, View) and _is_valid_target_view(elem):
                selected_views[_get_element_id_value(elem.Id)] = elem
        except Exception:
            pass

    if selected_views:
        return list(selected_views.values())

    if _is_valid_target_view(active_view):
        return [active_view]

    return []


target_views = _get_target_views_from_selection_or_active()
if not target_views:
    sys.exit(0)

views_to_hide_ids = {}
for target_view in target_views:
    section_ids = []
    collector = FilteredElementCollector(doc, target_view.Id).WhereElementIsNotElementType()
    for element in collector:
        if _is_section_view_element(element):
            if _is_own_section_reference(element, target_view):
                continue
            section_ids.append(element.Id)
    if section_ids:
        views_to_hide_ids[_get_element_id_value(target_view.Id)] = section_ids

if not views_to_hide_ids:
    sys.exit(0)

for target_view in target_views:
    if target_view.Id == active_view.Id:
        _zoom_extents()
        break

with revit.Transaction('Hide Section Views in Selected Views'):
    for target_view in target_views:
        section_ids = views_to_hide_ids.get(_get_element_id_value(target_view.Id))
        if not section_ids:
            continue
        target_view.HideElements(List[ElementId](section_ids))
