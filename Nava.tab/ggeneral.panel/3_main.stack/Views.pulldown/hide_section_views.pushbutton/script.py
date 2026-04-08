# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.DB import FilteredElementCollector, ElementId, View
from Autodesk.Revit.UI import RevitCommandId, PostableCommand
from System.Collections.Generic import List
from pyrevit import revit, forms, script
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


def _zoom_extents():
    try:
        cmd_id = RevitCommandId.LookupPostableCommandId(PostableCommand.ZoomToFit)
        uidoc.Application.PostCommand(cmd_id)
    except Exception:
        pass


def _pick_target_views(document):
    display_to_view = {}
    for view in FilteredElementCollector(document).OfClass(View).ToElements():
        try:
            if view.IsTemplate:
                continue
        except Exception:
            continue

        try:
            if str(view.ViewType) == 'DrawingSheet':
                continue
        except Exception:
            pass

        name = getattr(view, 'Name', None)
        if not name:
            continue

        label = '{} [{}]'.format(name, view.ViewType)
        display_to_view[label] = view

    if not display_to_view:
        return None

    selected = forms.SelectFromList.show(
        sorted(display_to_view.keys(), key=lambda x: x.lower()),
        title='Select Views',
        multiselect=True,
        button_name='Use Views'
    )

    if not selected:
        return []

    selected_views = []
    for selected_name in selected:
        view = display_to_view.get(selected_name)
        if view:
            selected_views.append(view)

    return selected_views


target_views = _pick_target_views(doc)
if not target_views:
    sys.exit(0)

views_to_hide_ids = {}
for target_view in target_views:
    section_ids = []
    collector = FilteredElementCollector(doc, target_view.Id).WhereElementIsNotElementType()
    for element in collector:
        if _is_section_view_element(element):
            section_ids.append(element.Id)
    if section_ids:
        views_to_hide_ids[target_view.Id.IntegerValue] = section_ids

if not views_to_hide_ids:
    sys.exit(0)

for target_view in target_views:
    if target_view.Id == active_view.Id:
        _zoom_extents()
        break

with revit.Transaction('Hide Section Views in Selected Views'):
    for target_view in target_views:
        section_ids = views_to_hide_ids.get(target_view.Id.IntegerValue)
        if not section_ids:
            continue
        target_view.HideElements(List[ElementId](section_ids))
