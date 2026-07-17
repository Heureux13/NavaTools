# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import DB, forms, revit, script

# Button info
# ======================================================================
__title__ = 'clean viewport'
__doc__ = '''
Moves north and west grid bubbles  1/4 inch away from view crop
removes south and east bubble and sets line ends to view crop edge 
'''

# Variables
# ======================================================================

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()
HALF_INCH_PAPER_FT = 0.5 / 12.00
EPS = 1e-9


def element_id_int(element_id):
    try:
        return element_id.IntegerValue
    except Exception:
        return int(element_id.Value)


def get_view_scale(view):
    scale_param = view.get_Parameter(DB.BuiltInParameter.VIEW_SCALE)
    if scale_param is not None:
        try:
            scale_value = float(scale_param.AsInteger())
            if scale_value > 0:
                return scale_value
        except Exception:
            pass

    try:
        scale_value = float(view.Scale)
        if scale_value > 0:
            return scale_value
    except Exception:
        pass

    return 1.0


def get_selected_target_views():
    target_by_id = {}
    selected_ids = list(uidoc.Selection.GetElementIds())

    for selected_id in selected_ids:
        element = doc.GetElement(selected_id)
        if element is None:
            continue

        if isinstance(element, DB.Viewport):
            view = doc.GetElement(element.ViewId)
            if view is not None:
                target_by_id[element_id_int(view.Id)] = view
            continue

        if isinstance(element, DB.View):
            target_by_id[element_id_int(element.Id)] = element

    if target_by_id:
        return list(target_by_id.values())

    active_view = doc.ActiveView
    if active_view is not None:
        return [active_view]

    return []


def is_supported_view(view):
    if view is None:
        return False

    if view.IsTemplate:
        return False

    if view.ViewType == DB.ViewType.DrawingSheet:
        return False

    if not view.CropBoxActive:
        return False

    return True


def get_visible_grids(view):
    return list(
        DB.FilteredElementCollector(doc, view.Id)
        .OfClass(DB.Grid)
        .WhereElementIsNotElementType()
        .ToElements()
    )


def get_line_curve_in_view(grid, view):
    curves = list(grid.GetCurvesInView(DB.DatumExtentType.ViewSpecific, view))
    if not curves:
        curves = list(grid.GetCurvesInView(DB.DatumExtentType.Model, view))
    if not curves:
        return None

    line = curves[0]
    if not isinstance(line, DB.Line):
        return None
    return line


def clip_line_to_rect(mid_x, mid_y, dir_x, dir_y, min_x, max_x, min_y, max_y):
    t_min = -1e30
    t_max = 1e30

    if abs(dir_x) < EPS:
        if mid_x < min_x or mid_x > max_x:
            return None
    else:
        tx1 = (min_x - mid_x) / dir_x
        tx2 = (max_x - mid_x) / dir_x
        t_min = max(t_min, min(tx1, tx2))
        t_max = min(t_max, max(tx1, tx2))

    if abs(dir_y) < EPS:
        if mid_y < min_y or mid_y > max_y:
            return None
    else:
        ty1 = (min_y - mid_y) / dir_y
        ty2 = (max_y - mid_y) / dir_y
        t_min = max(t_min, min(ty1, ty2))
        t_max = min(t_max, max(ty1, ty2))

    if t_max - t_min < EPS:
        return None

    return t_min, t_max


def build_outside_grid_line(view, line, offset):
    crop_box = view.CropBox
    to_local = crop_box.Transform.Inverse
    to_world = crop_box.Transform

    p0_local = to_local.OfPoint(line.GetEndPoint(0))
    p1_local = to_local.OfPoint(line.GetEndPoint(1))

    dir_x = p1_local.X - p0_local.X
    dir_y = p1_local.Y - p0_local.Y
    dir_len = (dir_x * dir_x + dir_y * dir_y) ** 0.5
    if dir_len < EPS:
        return None

    dir_x = dir_x / dir_len
    dir_y = dir_y / dir_len

    mid_x = (p0_local.X + p1_local.X) * 0.5
    mid_y = (p0_local.Y + p1_local.Y) * 0.5
    mid_z = (p0_local.Z + p1_local.Z) * 0.5

    # Asymmetric bounds: top/left go outside, right/bottom stay at crop border.
    min_x = crop_box.Min.X - offset
    max_x = crop_box.Max.X
    min_y = crop_box.Min.Y
    max_y = crop_box.Max.Y + offset
    if min_x >= max_x or min_y >= max_y:
        return None

    clipped = clip_line_to_rect(
        mid_x, mid_y, dir_x, dir_y, min_x, max_x, min_y, max_y)
    if clipped is None:
        return None

    t0, t1 = clipped
    new_start_local = DB.XYZ(mid_x + t0 * dir_x, mid_y + t0 * dir_y, mid_z)
    new_end_local = DB.XYZ(mid_x + t1 * dir_x, mid_y + t1 * dir_y, mid_z)

    new_start_world = to_world.OfPoint(new_start_local)
    new_end_world = to_world.OfPoint(new_end_local)
    return DB.Line.CreateBound(new_start_world, new_end_world)


def show_only_top_left_bubbles(view, grid, line, offset):
    crop_box = view.CropBox
    to_local = crop_box.Transform.Inverse

    min_x = crop_box.Min.X - offset
    max_y = crop_box.Max.Y + offset
    tol = max(1e-6, offset * 0.01)

    for datum_end, point_index in ((DB.DatumEnds.End0, 0), (DB.DatumEnds.End1, 1)):
        end_local = to_local.OfPoint(line.GetEndPoint(point_index))
        is_left = abs(end_local.X - min_x) <= tol
        is_top = abs(end_local.Y - max_y) <= tol
        should_show = is_left or is_top
        try:
            if should_show:
                grid.ShowBubbleInView(datum_end, view)
            else:
                grid.HideBubbleInView(datum_end, view)
        except Exception:
            pass


target_views = [v for v in get_selected_target_views() if is_supported_view(v)]
if not target_views:
    forms.alert(
        'Select one or more viewports (or views) with active crop regions, then run again.',
        title='Clean Viewport',
    )
    script.exit()

updated_grids = 0
updated_views = 0
skipped_views = []

with revit.Transaction('Set Grid Bubbles Outside (Top/Left Only)'):
    for view in target_views:
        offset_model_ft = HALF_INCH_PAPER_FT * get_view_scale(view)
        grids = get_visible_grids(view)
        if not grids:
            skipped_views.append(view.Name)
            continue

        updated_in_view = 0
        for grid in grids:
            line = get_line_curve_in_view(grid, view)
            if line is None:
                continue

            new_line = build_outside_grid_line(view, line, offset_model_ft)
            if new_line is None:
                continue

            grid.SetDatumExtentType(
                DB.DatumEnds.End0, view, DB.DatumExtentType.ViewSpecific)
            grid.SetDatumExtentType(
                DB.DatumEnds.End1, view, DB.DatumExtentType.ViewSpecific)
            grid.SetCurveInView(
                DB.DatumExtentType.ViewSpecific, view, new_line)
            show_only_top_left_bubbles(view, grid, new_line, offset_model_ft)
            updated_in_view += 1

        if updated_in_view > 0:
            updated_views += 1
            updated_grids += updated_in_view

# output.print_md(
#     'Updated {} grid(s) across {} view(s); top/left ends are {}" outside, right/bottom are flush, and bubbles are top/left only.'.format(
#         updated_grids, updated_views, 0.5
#     )
# )
if skipped_views:
    output.print_md('Views with no visible grids: {}'.format(
        ', '.join(skipped_views)))
