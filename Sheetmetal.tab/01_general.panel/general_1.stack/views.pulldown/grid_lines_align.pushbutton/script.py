# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import revit, script, forms, DB

# Button info
# ======================================================================
__title__ = 'Align Grid Lines'
__doc__ = '''
In the active section/elevation view, align visible grid extents so:
- bottom is flush with view bottom
- top is 1'-0" above view top
'''

# Variables
# ======================================================================
doc = revit.doc
uidoc = revit.uidoc
view = doc.ActiveView
output = script.get_output()
ONE_FOOT_FT = 1.0


def _is_supported_view(active_view):
    vt = active_view.ViewType
    return vt == DB.ViewType.Section or vt == DB.ViewType.Elevation


def _get_visible_grids(active_view):
    return list(
        DB.FilteredElementCollector(doc, active_view.Id)
        .OfClass(DB.Grid)
        .WhereElementIsNotElementType()
        .ToElements()
    )


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


def _get_target_views_from_selection():
    selected_ids = list(uidoc.Selection.GetElementIds())
    if not selected_ids:
        return []

    views_by_id = {}
    for element_id in selected_ids:
        elem = doc.GetElement(element_id)
        if not elem:
            continue

        try:
            if isinstance(elem, DB.Viewport):
                placed_view = doc.GetElement(elem.ViewId)
                if placed_view:
                    views_by_id[_get_element_id_value(placed_view.Id)] = placed_view
                continue
        except Exception:
            pass

        try:
            if isinstance(elem, DB.View):
                views_by_id[_get_element_id_value(elem.Id)] = elem
        except Exception:
            pass

    return list(views_by_id.values())


def _make_vertical_grid_line_in_view(active_view, grid, bottom_y, top_y):
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

    crop_t = active_view.CropBox.Transform
    to_local = crop_t.Inverse

    p0_local = to_local.OfPoint(line.GetEndPoint(0))
    p1_local = to_local.OfPoint(line.GetEndPoint(1))

    x = (p0_local.X + p1_local.X) / 2.0
    z = (p0_local.Z + p1_local.Z) / 2.0

    bottom_local = DB.XYZ(x, bottom_y, z)
    top_local = DB.XYZ(x, top_y, z)

    p0_world = crop_t.OfPoint(bottom_local)
    p1_world = crop_t.OfPoint(top_local)
    return DB.Line.CreateBound(p0_world, p1_world)


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


selected_views = _get_target_views_from_selection()
if selected_views:
    candidate_views = selected_views
else:
    candidate_views = [view]

target_views = []
for candidate in candidate_views:
    if _is_supported_view(candidate):
        target_views.append(candidate)

if not target_views:
    forms.alert(
        'Select section/elevation views (or viewports on a sheet), or open one and run again.',
        title='Align Grid Lines')
    script.exit()

updated_total = 0
updated_views = 0
skipped_no_crop = []
skipped_no_grids = []

with revit.Transaction('Align Grid Lines to Section Views'):
    for target_view in target_views:
        if not target_view.CropBoxActive:
            skipped_no_crop.append(target_view.Name)
            continue

        grids = _get_visible_grids(target_view)
        if not grids:
            skipped_no_grids.append(target_view.Name)
            continue

        crop = target_view.CropBox
        bottom_y = crop.Min.Y
        top_y = crop.Max.Y + ONE_FOOT_FT

        updated_in_view = 0
        for grid in grids:
            new_line = _make_vertical_grid_line_in_view(target_view, grid, bottom_y, top_y)
            if not new_line:
                continue
            try:
                _set_grid_view_specific_extent(target_view, grid, new_line)
                updated_in_view += 1
            except Exception:
                pass

        if updated_in_view > 0:
            updated_views += 1
            updated_total += updated_in_view

output.print_md('Aligned {} grid(s) across {} view(s).'.format(updated_total, updated_views))
if skipped_no_crop:
    output.print_md('Skipped (crop off): {}'.format(', '.join(skipped_no_crop)))
if skipped_no_grids:
    output.print_md('Skipped (no grids): {}'.format(', '.join(skipped_no_grids)))
