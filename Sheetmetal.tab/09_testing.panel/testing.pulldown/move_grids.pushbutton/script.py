# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import revit, script, forms, DB
from Autodesk.Revit.DB import Viewport, ViewSheet

# Button info
# ======================================================================
__title__ = 'Move Grids'
__doc__ = '''
Select one or more viewports on a sheet, or open a section/elevation view.
Moves visible grid bubbles so the grid ends sit 1/2" inside the view border.
'''

# Variables
# ======================================================================

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()
HALF_INCH_FT = 0.5 / 12.0


def _is_supported_view(active_view):
    try:
        vt = active_view.ViewType
    except Exception:
        return False

    return vt in (
        DB.ViewType.Section,
        DB.ViewType.Elevation,
        DB.ViewType.FloorPlan,
        DB.ViewType.CeilingPlan,
        DB.ViewType.EngineeringPlan,
        DB.ViewType.AreaPlan,
    )


def _get_visible_grids(active_view):
    grids = list(
        DB.FilteredElementCollector(doc, active_view.Id)
        .OfCategory(DB.BuiltInCategory.OST_Grids)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    if grids:
        return grids

    return list(
        DB.FilteredElementCollector(doc, active_view.Id)
        .OfClass(DB.Grid)
        .WhereElementIsNotElementType()
        .ToElements()
    )


def _get_target_views_from_selection():
    selected_ids = list(uidoc.Selection.GetElementIds())
    target_views = []

    for element_id in selected_ids:
        elem = doc.GetElement(element_id)
        if not elem:
            continue

        if isinstance(elem, Viewport):
            placed_view = doc.GetElement(elem.ViewId)
            if placed_view and _is_supported_view(placed_view):
                target_views.append(placed_view)
            continue

        if isinstance(elem, DB.View) and _is_supported_view(elem):
            target_views.append(elem)

    unique_views = []
    seen_ids = set()
    for target_view in target_views:
        view_id = target_view.Id.IntegerValue if hasattr(target_view.Id, 'IntegerValue') else int(target_view.Id)
        if view_id in seen_ids:
            continue
        seen_ids.add(view_id)
        unique_views.append(target_view)

    return unique_views


def _make_grid_line_in_view(active_view, grid):
    curves = []
    try:
        curves = list(grid.GetCurvesInView(DB.DatumExtentType.ViewSpecific, active_view))
    except Exception:
        pass

    if not curves:
        try:
            curves = list(grid.GetCurvesInView(DB.DatumExtentType.Model, active_view))
        except Exception:
            return None

    if not curves:
        return None

    line = curves[0]
    if not isinstance(line, DB.Line):
        return None

    crop_box = active_view.CropBox
    crop_t = crop_box.Transform
    to_local = crop_t.Inverse

    p0_local = to_local.OfPoint(line.GetEndPoint(0))
    p1_local = to_local.OfPoint(line.GetEndPoint(1))

    direction = p1_local - p0_local
    if direction.GetLength() <= 1e-9:
        return None

    sheet_offset_ft = HALF_INCH_FT * float(getattr(active_view, 'Scale', 1) or 1)

    axis_values = [abs(direction.X), abs(direction.Y), abs(direction.Z)]
    line_axis = axis_values.index(max(axis_values))

    min_vals = [crop_box.Min.X, crop_box.Min.Y, crop_box.Min.Z]
    max_vals = [crop_box.Max.X, crop_box.Max.Y, crop_box.Max.Z]

    # Keep non-bubble ends fixed. Only bubble ends are snapped to a fixed
    # offset outside the crop boundary so reruns are stable.
    center_axis = (min_vals[line_axis] + max_vals[line_axis]) / 2.0

    end0_vals = [p0_local.X, p0_local.Y, p0_local.Z]
    end1_vals = [p1_local.X, p1_local.Y, p1_local.Z]

    try:
        end0_has_bubble = bool(grid.IsBubbleVisibleInView(DB.DatumEnds.End0, active_view))
    except Exception:
        end0_has_bubble = False

    try:
        end1_has_bubble = bool(grid.IsBubbleVisibleInView(DB.DatumEnds.End1, active_view))
    except Exception:
        end1_has_bubble = False

    if not end0_has_bubble and not end1_has_bubble:
        return None

    if end0_has_bubble:
        if end0_vals[line_axis] >= center_axis:
            end0_vals[line_axis] = max_vals[line_axis] + sheet_offset_ft
        else:
            end0_vals[line_axis] = min_vals[line_axis] - sheet_offset_ft

    if end1_has_bubble:
        if end1_vals[line_axis] >= center_axis:
            end1_vals[line_axis] = max_vals[line_axis] + sheet_offset_ft
        else:
            end1_vals[line_axis] = min_vals[line_axis] - sheet_offset_ft

    new_p0 = DB.XYZ(end0_vals[0], end0_vals[1], end0_vals[2])
    new_p1 = DB.XYZ(end1_vals[0], end1_vals[1], end1_vals[2])
    if new_p0.DistanceTo(new_p1) <= 1e-9:
        return None

    return DB.Line.CreateBound(crop_t.OfPoint(new_p0), crop_t.OfPoint(new_p1))


def _set_grid_view_specific_extent(active_view, grid, new_line):
    try:
        grid.SetDatumExtentType(DB.DatumEnds.End0, active_view, DB.DatumExtentType.ViewSpecific)
    except Exception:
        pass

    try:
        grid.SetDatumExtentType(DB.DatumEnds.End1, active_view, DB.DatumExtentType.ViewSpecific)
    except Exception:
        pass

    grid.SetCurveInView(DB.DatumExtentType.ViewSpecific, active_view, new_line)


def _move_grid_with_pin_restore(active_view, grid, new_line):
    was_pinned = False
    try:
        was_pinned = bool(grid.Pinned)
    except Exception:
        was_pinned = False

    if was_pinned:
        try:
            grid.Pinned = False
        except Exception:
            return False

    try:
        doc.Regenerate()
    except Exception:
        pass

    try:
        _set_grid_view_specific_extent(active_view, grid, new_line)
    except Exception:
        return False

    try:
        doc.Regenerate()
    except Exception:
        pass

    return True


target_views = _get_target_views_from_selection()
if not target_views:
    active_view = doc.ActiveView
    if isinstance(active_view, ViewSheet):
        forms.alert(
            'Select one or more section/elevation/plan viewports on the sheet, then run again.',
            title='Move Grids'
        )
        script.exit()

    if _is_supported_view(active_view):
        target_views = [active_view]

if not target_views:
    forms.alert(
        'Select one or more section/elevation/plan viewports on a sheet, or open one and run again.',
        title='Move Grids'
    )
    script.exit()

updated_total = 0
updated_views = 0
skipped_no_crop = []
skipped_no_grids = []

with revit.Transaction('Move Grid Bubbles 1/2 Inch In'):
    for target_view in target_views:
        if not target_view.CropBoxActive:
            skipped_no_crop.append(target_view.Name)
            continue

        grids = _get_visible_grids(target_view)
        if not grids:
            skipped_no_grids.append(target_view.Name)
            continue

        updated_in_view = 0
        for grid in grids:
            new_line = _make_grid_line_in_view(target_view, grid)
            if not new_line:
                continue
            try:
                if _move_grid_with_pin_restore(target_view, grid, new_line):
                    updated_in_view += 1
            except Exception:
                pass

        if updated_in_view > 0:
            updated_views += 1
            updated_total += updated_in_view

# output.print_md('Moved {} grid(s) across {} view(s).'.format(updated_total, updated_views))
if skipped_no_crop:
    output.print_md('Skipped (crop off or invalid crop): {}'.format(', '.join(skipped_no_crop)))
if skipped_no_grids:
    output.print_md('Skipped (no grids): {}'.format(', '.join(skipped_no_grids)))
