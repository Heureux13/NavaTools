# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import forms, revit, script
from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    CopyPasteOptions,
    ElementId,
    ElementTransformUtils,
    FilteredElementCollector,
    Transform,
    View,
    ViewSheet,
)
from System.Collections.Generic import List

# Button info
# ======================================================================
__title__ = 'Room Tags Copy'
__doc__ = '''
Select all matching Room Tags in active view,
then copy/paste them into selected views.
'''

# Variables
# ======================================================================

doc = revit.doc
uidoc = revit.uidoc
source_view = revit.active_view
output = script.get_output()


class ListOption(object):
    def __init__(self, item, display_name):
        self.item = item
        self.display_name = display_name


def get_view_display_name(view):
    try:
        return '{} | {}'.format(view.ViewType, view.Name)
    except Exception:
        return str(view.Id)


def get_tag_family_and_type(tag):
    family_name = None
    type_name = None

    try:
        tag_type = doc.GetElement(tag.GetTypeId())
    except Exception:
        tag_type = None

    if tag_type:
        try:
            family_name = getattr(tag_type, 'FamilyName', None)
        except Exception:
            family_name = None

        try:
            type_name = getattr(tag_type, 'Name', None)
        except Exception:
            type_name = None

        if not family_name:
            try:
                fam_param = tag_type.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
                if fam_param:
                    family_name = fam_param.AsString()
            except Exception:
                pass

        if not type_name:
            try:
                name_param = tag_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                if name_param:
                    type_name = name_param.AsString()
            except Exception:
                pass

    return family_name, type_name


def collect_target_views(document, active):
    all_views = [
        v for v in FilteredElementCollector(document).OfClass(View)
        if not v.IsTemplate and not isinstance(v, ViewSheet)
    ]

    candidates = [
        v for v in all_views
        if v.Id != active.Id and v.ViewType == active.ViewType
    ]

    return sorted(candidates, key=lambda v: get_view_display_name(v).lower())


def collect_annotation_options(tags):
    grouped = {}
    for tag in tags:
        family_name, type_name = get_tag_family_and_type(tag)
        family_name = family_name or '(Unknown Family)'
        type_name = type_name or '(Unknown Type)'
        key = (family_name, type_name)
        if key not in grouped:
            grouped[key] = []
        grouped[key].append(tag)

    options = []
    for key in sorted(grouped.keys(), key=lambda k: '{}:{}'.format(k[0], k[1]).lower()):
        family_name, type_name = key
        label = '{} : {} ({})'.format(family_name, type_name, len(grouped[key]))
        options.append(ListOption(key, label))

    return options, grouped


try:
    room_tags = list(
        FilteredElementCollector(doc, source_view.Id)
        .OfCategory(BuiltInCategory.OST_RoomTags)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    room_tags = [t for t in room_tags if t.OwnerViewId == source_view.Id]

    if not room_tags:
        output.print_md('## No room tags found in the active view.')
        script.exit()

    annotation_options, grouped_tags = collect_annotation_options(room_tags)
    picked_annotations = forms.SelectFromList.show(
        annotation_options,
        name_attr='display_name',
        multiselect=True,
        title='Select Annotation Name(s) to Copy'
    )
    if not picked_annotations:
        script.exit()

    selected_keys = set([opt.item for opt in picked_annotations])

    matching_tags = []
    for tag in room_tags:
        family_name, type_name = get_tag_family_and_type(tag)
        key = (family_name or '(Unknown Family)', type_name or '(Unknown Type)')
        if key in selected_keys:
            matching_tags.append(tag)

    if not matching_tags:
        output.print_md('## No matching room tags found for the selected annotation name(s).')
        script.exit()

    selected_ids = List[ElementId]([tag.Id for tag in matching_tags])
    uidoc.Selection.SetElementIds(selected_ids)

    target_views = collect_target_views(doc, source_view)
    if not target_views:
        output.print_md('## No compatible target views found.')
        script.exit()

    options = [ListOption(v, get_view_display_name(v)) for v in target_views]
    picked = forms.SelectFromList.show(
        options,
        name_attr='display_name',
        multiselect=True,
        title='Select Views to Paste Room Tags'
    )
    if not picked:
        script.exit()

    destination_views = [opt.item for opt in picked]
    copied_count = 0
    failed = []

    with revit.Transaction('Copy Room Tags To Views'):
        for destination_view in destination_views:
            try:
                ElementTransformUtils.CopyElements(
                    source_view,
                    selected_ids,
                    destination_view,
                    Transform.Identity,
                    CopyPasteOptions()
                )
                copied_count += 1
            except Exception as copy_error:
                failed.append((get_view_display_name(destination_view), str(copy_error)))

    output.print_md(
        '### Selected {} room tags and pasted to {} view(s).'.format(
            len(matching_tags),
            copied_count,
        )
    )

    if failed:
        output.print_md('### Failed views:')
        for view_name, err in failed:
            output.print_md('- {} | {}'.format(view_name, err))

except Exception as err:
    output.print_md('## Error: {}'.format(str(err)))
    script.exit()
