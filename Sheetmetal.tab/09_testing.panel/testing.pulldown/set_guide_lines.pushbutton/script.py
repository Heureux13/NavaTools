# -*- coding: utf-8 -*-
# pyright: reportAttributeAccessIssue=false, reportCallIssue=false, reportGeneralTypeIssues=false, reportUndefinedVariable=false, reportInvalidTypeArguments=false, reportMissingImports=false
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import script, DB, forms
from System.Collections.Generic import List

# Button info
# ======================================================================
__title__ = 'Select Ducts in Scope Box'
__doc__ = '''
Select MEP ducts inside selected scope boxes.

Usage:
1. Select one or more Scope Boxes and run the tool
2. If none are selected, pick Scope Boxes from a list
3. If no Scope Box is chosen, tool uses active 3D view Section Box
4. All intersecting ducts are selected
'''

# Variables
# ======================================================================

output = script.get_output()

try:
    _revit_app = __revit__
except NameError:
    _revit_app = None

uidoc = _revit_app.ActiveUIDocument if _revit_app else None
if uidoc is None:
    output.print_md('## Select Ducts in Box')
    output.print_md('This command must run inside Revit.')
    raise RuntimeError('No ActiveUIDocument. Run inside Revit.')

doc = uidoc.Document
view = doc.ActiveView


def bbox_to_outline(bbox):
    if not bbox:
        return None

    transform = bbox.Transform if bbox.Transform else DB.Transform.Identity
    min_pt = bbox.Min
    max_pt = bbox.Max

    corners_local = [
        DB.XYZ(min_pt.X, min_pt.Y, min_pt.Z),
        DB.XYZ(min_pt.X, min_pt.Y, max_pt.Z),
        DB.XYZ(min_pt.X, max_pt.Y, min_pt.Z),
        DB.XYZ(min_pt.X, max_pt.Y, max_pt.Z),
        DB.XYZ(max_pt.X, min_pt.Y, min_pt.Z),
        DB.XYZ(max_pt.X, min_pt.Y, max_pt.Z),
        DB.XYZ(max_pt.X, max_pt.Y, min_pt.Z),
        DB.XYZ(max_pt.X, max_pt.Y, max_pt.Z),
    ]
    corners_world = [transform.OfPoint(pt) for pt in corners_local]

    xs = [pt.X for pt in corners_world]
    ys = [pt.Y for pt in corners_world]
    zs = [pt.Z for pt in corners_world]

    world_min = DB.XYZ(min(xs), min(ys), min(zs))
    world_max = DB.XYZ(max(xs), max(ys), max(zs))
    return DB.Outline(world_min, world_max)


def collect_selected_scope_boxes(document, selection_ids):
    boxes = []
    for eid in selection_ids:
        elem = document.GetElement(eid)
        if not elem or not elem.Category:
            continue
        if elem.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_VolumeOfInterest):
            boxes.append(elem)
    return boxes


def get_scope_box_bbox(scope_box, active_view):
    """Try multiple access patterns because scope boxes can return None in some contexts."""
    bbox = None
    try:
        bbox = scope_box.get_BoundingBox(None)
    except Exception:
        bbox = None

    if bbox:
        return bbox

    try:
        bbox = scope_box.get_BoundingBox(active_view)
    except Exception:
        bbox = None

    return bbox


def prompt_for_scope_boxes(document):
    scope_boxes = list(
        DB.FilteredElementCollector(document)
        # type: ignore[arg-type]
        .OfCategory(DB.BuiltInCategory.OST_VolumeOfInterest)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    if not scope_boxes:
        return []

    scope_boxes = sorted(scope_boxes, key=lambda box: (
        box.Name, box.Id.IntegerValue))
    options = ['{} | ID {}'.format(box.Name, box.Id.IntegerValue)
               for box in scope_boxes]

    selected_options = forms.SelectFromList.show(
        options,
        title='Select Scope Boxes',
        multiselect=True,
        button_name='Use Scope Boxes'
    )
    if not selected_options:
        return []

    return [
        box for box, option in zip(scope_boxes, options)
        if option in selected_options
    ]


def get_target_outlines(document, active_view, selected_ids):
    outlines = []
    source_label = ''

    scope_boxes = collect_selected_scope_boxes(document, selected_ids)
    if not scope_boxes:
        scope_boxes = prompt_for_scope_boxes(document)

    if scope_boxes:
        for box in scope_boxes:
            bbox = get_scope_box_bbox(box, active_view)
            outline = bbox_to_outline(bbox)
            if outline:
                outlines.append(outline)
        source_label = 'scope box selection'
        return outlines, source_label

    if isinstance(active_view, DB.View3D) and active_view.IsSectionBoxActive:
        section_bbox = active_view.GetSectionBox()
        outline = bbox_to_outline(section_bbox)
        if outline:
            outlines.append(outline)
            source_label = 'active 3D section box'

    return outlines, source_label


def collect_ducts_in_outlines(document, active_view, outlines):
    if not outlines:
        return []

    category_ids = List[DB.ElementId]([
        DB.ElementId(DB.BuiltInCategory.OST_DuctCurves),
        DB.ElementId(DB.BuiltInCategory.OST_FlexDuctCurves),
        DB.ElementId(DB.BuiltInCategory.OST_DuctFitting),
        DB.ElementId(DB.BuiltInCategory.OST_DuctAccessory),
        DB.ElementId(DB.BuiltInCategory.OST_FabricationDuctwork),
    ])
    cat_filter = DB.ElementMulticategoryFilter(category_ids)

    collector = (
        DB.FilteredElementCollector(document, active_view.Id)
        .WherePasses(cat_filter)
        .WhereElementIsNotElementType()
    )

    def outlines_overlap(outline_a, outline_b):
        if not outline_a or not outline_b:
            return False

        min_a = outline_a.MinimumPoint
        max_a = outline_a.MaximumPoint
        min_b = outline_b.MinimumPoint
        max_b = outline_b.MaximumPoint

        return not (
            max_a.X < min_b.X or max_b.X < min_a.X or
            max_a.Y < min_b.Y or max_b.Y < min_a.Y or
            max_a.Z < min_b.Z or max_b.Z < min_a.Z
        )

    found = {}
    for elem in collector.ToElements():
        elem_bbox = elem.get_BoundingBox(
            active_view) or elem.get_BoundingBox(None)
        elem_outline = bbox_to_outline(elem_bbox)
        if not elem_outline:
            continue

        for scope_outline in outlines:
            if outlines_overlap(elem_outline, scope_outline):
                found[elem.Id.IntegerValue] = elem.Id
                break

    return list(found.values())


selection_ids = uidoc.Selection.GetElementIds()
target_outlines, source = get_target_outlines(doc, view, selection_ids)

if not target_outlines:
    output.print_md('## Select Ducts in Box')
    output.print_md(
        'No valid target found. Select one or more Scope Boxes, or run in a 3D view with Section Box active.'
    )
    script.exit()

duct_ids = collect_ducts_in_outlines(doc, view, target_outlines)

if not duct_ids:
    output.print_md('## Select Ducts in Box')
    output.print_md('No ducts found in the {}.'.format(source))
    script.exit()

uidoc.Selection.SetElementIds(List[DB.ElementId](duct_ids))

output.print_md('## Select Ducts in Box')
output.print_md('Selected **{}** ducts from {}.'.format(len(duct_ids), source))
