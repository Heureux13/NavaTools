# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from pyrevit import revit, forms, script
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    RevitLinkInstance,
    ElementId,
    BuiltInParameter,
    TemporaryViewMode,
)
from System.Collections.Generic import List
import sys
import json
import os

# Button info
# ===================================================
__title__ = "Hide Links"
__doc__ = """
Hide or unhide selected Revit links in the active view.
"""

# Variables
# ==================================================
doc = revit.doc
active_view = doc.ActiveView
output = script.get_output()

# Store hidden link IDs in view's temporary data
HIDDEN_KEY = 'NavaTools_HiddenLinks_{}'.format(active_view.Id.IntegerValue)


def get_hidden_link_ids():
    try:
        data_file = script.get_document_data_file(doc.PathName or 'unsaved', HIDDEN_KEY)
        with open(data_file, 'r') as f:
            data = json.load(f)
            return set(data.get('hidden_ids', []))
    except Exception:
        return set()


def save_hidden_link_ids(hidden_ids):
    try:
        data = {'hidden_ids': list(hidden_ids)}
        data_file = script.get_document_data_file(doc.PathName or 'unsaved', HIDDEN_KEY)
        with open(data_file, 'w') as f:
            json.dump(data, f)
    except Exception:
        pass


def is_hidden_in_view(view, element_id):
    # For older Revit APIs without IsTemporaryViewModeActive,
    # we can't reliably detect temporary hide state.
    # Check permanent hide only.
    try:
        element = doc.GetElement(element_id)
        if element and element.IsHidden(view):
            return True
    except Exception:
        pass
    try:
        if view.IsElementHidden(element_id):
            return True
    except Exception:
        pass
    try:
        if element_id in view.GetHiddenElementIds():
            return True
    except Exception:
        pass
    return False


def clear_temporary_hide(view):
    try:
        view.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate)
    except Exception:
        pass


def get_link_nickname(link):
    param = link.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
    if param and param.HasValue:
        nickname = param.AsString()
        if nickname:
            return nickname.strip()
    param = link.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
    if param and param.HasValue:
        nickname = param.AsString()
        if nickname:
            return nickname.strip()
    try:
        link_type = doc.GetElement(link.GetTypeId())
    except Exception:
        link_type = None
    if link_type:
        param = link_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS)
        if param and param.HasValue:
            nickname = param.AsString()
            if nickname:
                return nickname.strip()
        param = link_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_MARK)
        if param and param.HasValue:
            nickname = param.AsString()
            if nickname:
                return nickname.strip()
    return None


links_all = list(
    FilteredElementCollector(doc)
    .OfClass(RevitLinkInstance)
    .WhereElementIsNotElementType()
)

if not links_all:
    output.print_md('**No Revit links found in this model.**')
    sys.exit(0)

links_in_view = links_all

hidden_link_ids = get_hidden_link_ids()

display_map = {}
pad_width = max(2, len(str(len(links_in_view))))
for idx, link in enumerate(links_in_view, start=1):
    idx_label = str(idx).zfill(pad_width)
    is_hidden = link.Id.IntegerValue in hidden_link_ids
    status = u'✓ Hidden' if is_hidden else u'Visible'
    nickname = get_link_nickname(link)
    if nickname:
        display_name = u'{}  |  {}  |  {} — {} (Id: {})'.format(
            status, idx_label, nickname, link.Name, link.Id.IntegerValue
        )
    else:
        display_name = u'{}  |  {}  |  {} (Id: {})'.format(
            status, idx_label, link.Name, link.Id.IntegerValue
        )
    display_map[display_name] = link.Id

selected_names = forms.SelectFromList.show(
    sorted(display_map.keys()),
    title='Select Links to Toggle',
    multiselect=True,
    button_name='Toggle Selected'
)

if not selected_names:
    sys.exit(0)

selected_ids = set(display_map[name] for name in selected_names if name in display_map)

# Toggle hidden state for selected links
for link_id in selected_ids:
    link_id_int = link_id.IntegerValue
    if link_id_int in hidden_link_ids:
        hidden_link_ids.remove(link_id_int)
    else:
        hidden_link_ids.add(link_id_int)

# Build final hide list
ids_to_hide = List[ElementId]()
for link in links_in_view:
    if link.Id.IntegerValue in hidden_link_ids:
        ids_to_hide.Add(link.Id)

# Save state and apply
save_hidden_link_ids(hidden_link_ids)

with revit.Transaction('Toggle Link Visibility'):
    try:
        clear_temporary_hide(active_view)
        if ids_to_hide.Count > 0:
            active_view.HideElementsTemporary(ids_to_hide)
    except Exception:
        output.print_md('**This view does not allow temporary hiding of links.**')
