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

NICKNAME_KEY = 'NavaTools_LinkNicknames'


def get_all_nicknames():
    try:
        data_file = script.get_document_data_file(doc.PathName or 'unsaved', NICKNAME_KEY)
        if os.path.exists(data_file):
            with open(data_file, 'r') as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def save_all_nicknames(nicknames_dict):
    try:
        data_file = script.get_document_data_file(doc.PathName or 'unsaved', NICKNAME_KEY)
        with open(data_file, 'w') as f:
            json.dump(nicknames_dict, f)
    except Exception:
        pass


def get_link_nickname(link, nicknames_dict):
    # Check custom nicknames first
    link_id_str = str(link.Id.IntegerValue)
    if link_id_str in nicknames_dict:
        return nicknames_dict[link_id_str]
    return None


links_all = list(
    FilteredElementCollector(doc)
    .OfClass(RevitLinkInstance)
    .WhereElementIsNotElementType()
)

if not links_all:
    output.print_md('**No Revit links found in this model.**')
    sys.exit(0)

# Filter to only visible links (collector with view ID gets visible elements)
links_in_view = list(
    FilteredElementCollector(doc, active_view.Id)
    .OfClass(RevitLinkInstance)
    .WhereElementIsNotElementType()
)

nicknames_dict = get_all_nicknames()

display_map = {}
link_map = {}
pad_width = max(2, len(str(len(links_in_view))))
for idx, link in enumerate(links_in_view, start=1):
    idx_label = str(idx).zfill(pad_width)
    nickname = get_link_nickname(link, nicknames_dict)
    if nickname:
        display_name = u'{}  |  {} — {} (Id: {})'.format(
            idx_label, nickname, link.Name, link.Id.IntegerValue
        )
    else:
        display_name = u'{}  |  {} (Id: {})'.format(
            idx_label, link.Name, link.Id.IntegerValue
        )
    display_map[display_name] = link.Id
    link_map[display_name] = link

# Show selection dialog with option to set nicknames


class LinkSelectionForm(forms.WPFWindow):
    def __init__(self):
        pass


# Ask what user wants to do first
action = forms.CommandSwitchWindow.show(
    ['Hide Links', 'Rename Links'],
    message='What would you like to do?'
)

if not action:
    sys.exit(0)

if action == 'Rename Links':
    # Select links to rename
    selected_names = forms.SelectFromList.show(
        sorted(display_map.keys()),
        title='Select Links to Rename',
        multiselect=True,
        button_name='Rename Selected'
    )

    if not selected_names:
        sys.exit(0)

    # Rename each selected link
    for name in selected_names:
        if name in link_map:
            link = link_map[name]
            current_nick = get_link_nickname(link, nicknames_dict)
            new_nick = forms.ask_for_string(
                prompt='Enter nickname for: {}'.format(link.Name),
                default=current_nick or '',
                title='Set Link Nickname'
            )
            if new_nick:
                nicknames_dict[str(link.Id.IntegerValue)] = new_nick
            elif new_nick == '' and current_nick:
                # Clear nickname if empty string entered
                nicknames_dict.pop(str(link.Id.IntegerValue), None)
    save_all_nicknames(nicknames_dict)
    sys.exit(0)

# Hide Links workflow
selected_names = forms.SelectFromList.show(
    sorted(display_map.keys()),
    title='Select Links to Hide',
    multiselect=True,
    button_name='Hide Selected'
)

if not selected_names:
    sys.exit(0)

selected_ids = [display_map[name] for name in selected_names if name in display_map]

# Build hide list with selected links
ids_to_hide = List[ElementId]()
for link_id in selected_ids:
    ids_to_hide.Add(link_id)

with revit.Transaction('Hide Selected Links'):
    try:
        if ids_to_hide.Count > 0:
            active_view.HideElementsTemporary(ids_to_hide)
    except Exception:
        output.print_md('**This view does not allow temporary hiding of links.**')
