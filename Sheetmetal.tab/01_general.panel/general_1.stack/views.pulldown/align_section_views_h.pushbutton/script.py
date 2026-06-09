# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import revit, forms, script
from Autodesk.Revit.DB import (
    DatumExtentType,
    FilteredElementCollector,
    Level,
    Viewport,
    ViewSheet,
    XYZ,
)

# Button info
# ======================================================================
__title__ = 'Align Views Horizontally'
__doc__ = '''
Select 2+ section viewports on an active sheet, then run.
  - Aligns viewports vertically so their floor levels line up
  - Equally spaces viewports horizontally (outer bounds preserved)
  - Places every viewport title 0.5" below the largest view
'''

# Constants
# ======================================================================
HALF_INCH_FT = 0.15 / 12.0

# Variables
# ======================================================================
output = script.get_output()
doc = revit.doc
uidoc = revit.uidoc


def element_id_int(elem_id):
    """Return a stable int for ElementId across Revit API versions."""
    if hasattr(elem_id, 'IntegerValue'):
        return elem_id.IntegerValue
    if hasattr(elem_id, 'Value'):
        return elem_id.Value
    return int(str(elem_id))


# ── 1. Validate sheet is active ──────────────────────────────────────
active_view = revit.active_view
if not isinstance(active_view, ViewSheet):
    forms.alert(
        'Active view must be a sheet.\nOpen a sheet and try again.',
        exitscript=True
    )

# ── 2. Get selected viewports ────────────────────────────────────────
sel_ids = uidoc.Selection.GetElementIds()
viewports = []
for eid in sel_ids:
    el = doc.GetElement(eid)
    if isinstance(el, Viewport):
        viewports.append(el)

if len(viewports) < 2:
    forms.alert('Select at least 2 viewports on the sheet.', exitscript=True)

# ── 3. Gather crop data per viewport ─────────────────────────────────
vp_data = []
for vp in viewports:
    view = doc.GetElement(vp.ViewId)
    if view is None:
        continue
    try:
        scale = float(view.Scale)
        if scale <= 0:
            continue

        crop_box = view.CropBox
        transform = crop_box.Transform
        local_x_center = (crop_box.Min.X + crop_box.Max.X) / 2.0
        local_y_center = (crop_box.Min.Y + crop_box.Max.Y) / 2.0
        local_z_center = (crop_box.Min.Z + crop_box.Max.Z) / 2.0

        # Pick the local crop axis that best represents world elevation.
        # In most sections this is Y, but it can vary with transform orientation.
        z_contrib = [
            abs(transform.BasisX.Z),
            abs(transform.BasisY.Z),
            abs(transform.BasisZ.Z),
        ]
        vertical_axis = z_contrib.index(max(z_contrib))

        center_coords = [local_x_center, local_y_center, local_z_center]
        min_coords = [crop_box.Min.X, crop_box.Min.Y, crop_box.Min.Z]
        max_coords = [crop_box.Max.X, crop_box.Max.Y, crop_box.Max.Z]

        # World Z (elevation) at crop-box center
        local_mid = XYZ(center_coords[0], center_coords[1], center_coords[2])
        world_mid = transform.OfPoint(local_mid)
        elev_center = world_mid.Z

        # Convert crop vertical extents to project elevation range
        bottom_coords = list(center_coords)
        top_coords = list(center_coords)
        bottom_coords[vertical_axis] = min_coords[vertical_axis]
        top_coords[vertical_axis] = max_coords[vertical_axis]

        world_bottom = transform.OfPoint(XYZ(bottom_coords[0], bottom_coords[1], bottom_coords[2]))
        world_top = transform.OfPoint(XYZ(top_coords[0], top_coords[1], top_coords[2]))
        elev_bottom = min(world_bottom.Z, world_top.Z)
        elev_top = max(world_bottom.Z, world_top.Z)

        box_outline = vp.GetBoxOutline()
        box_center = vp.GetBoxCenter()
        box_w = box_outline.MaximumPoint.X - box_outline.MinimumPoint.X
        box_h = box_outline.MaximumPoint.Y - box_outline.MinimumPoint.Y

        vp_data.append({
            'vp': vp,
            'view': view,
            'scale': scale,
            'elev_center': elev_center,
            'elev_bottom': elev_bottom,
            'elev_top': elev_top,
            'box_center': box_center,
            'box_w': box_w,
            'box_h': box_h,
        })
    except Exception as exc:
        output.print_md('Skipped viewport {}: {}'.format(element_id_int(vp.Id), exc))

if len(vp_data) < 2:
    forms.alert(
        'Need at least 2 valid section viewports with crop boxes.',
        exitscript=True
    )

# ── 4. Reference level: one level shared by all selected views ───────
avg_center_elev = sum(d['elev_center'] for d in vp_data) / len(vp_data)

visible_by_view = []
for d in vp_data:
    visible_levels = list(
        FilteredElementCollector(doc, d['view'].Id)
        .OfClass(Level)
        .ToElements()
    )
    visible_by_view.append(visible_levels)

common_levels = []
if visible_by_view:
    common_ids = set(element_id_int(lvl.Id) for lvl in visible_by_view[0])
    for visible in visible_by_view[1:]:
        common_ids &= set(element_id_int(lvl.Id) for lvl in visible)

    by_id = {}
    for levels in visible_by_view:
        for lvl in levels:
            by_id[element_id_int(lvl.Id)] = lvl
    common_levels = [by_id[i] for i in common_ids if i in by_id]

if not common_levels:
    forms.alert(
        'No common visible level found across all selected views.\n'
        'Select views that share at least one visible level and try again.',
        exitscript=True
    )

ref_level = min(common_levels, key=lambda lvl: abs(lvl.Elevation - avg_center_elev))

# ── 5. Level alignment – compute new V centers from level graphics ───
# Use the actual reference level line shown in each section view.
for d in vp_data:
    view = d['view']
    curves = []
    try:
        curves = list(ref_level.GetCurvesInView(DatumExtentType.ViewSpecific, view))
    except Exception:
        curves = []

    if not curves:
        try:
            curves = list(ref_level.GetCurvesInView(DatumExtentType.Model, view))
        except Exception:
            curves = []

    if not curves:
        forms.alert(
            'Reference level "{}" has no curve in one of the selected views.\n'
            'Check level extents/visibility and try again.'.format(ref_level.Name),
            exitscript=True
        )

    level_curve = curves[0]
    level_mid = level_curve.Evaluate(0.5, True)

    crop_box = view.CropBox
    inv = crop_box.Transform.Inverse
    local_pt = inv.OfPoint(level_mid)
    local_center_v = (crop_box.Min.Y + crop_box.Max.Y) / 2.0
    model_offset_v = local_pt.Y - local_center_v

    d['anchor_level'] = ref_level
    d['ref_v'] = d['box_center'].Y + model_offset_v / d['scale']
    d['anchor_offset_v'] = model_offset_v

# Align all to the average (minimises total displacement)
avg_ref_v = sum(d['ref_v'] for d in vp_data) / len(vp_data)

for d in vp_data:
    delta = d['anchor_offset_v'] / d['scale']
    d['new_v'] = avg_ref_v - delta

# ── 6. Equal horizontal spacing – compute new U centers ──────────────
# Sort left-to-right by current box center
vp_sorted = sorted(vp_data, key=lambda d: d['box_center'].X)

left_bound = vp_sorted[0]['box_center'].X - vp_sorted[0]['box_w'] / 2.0
right_bound = vp_sorted[-1]['box_center'].X + vp_sorted[-1]['box_w'] / 2.0
total_span = right_bound - left_bound
sum_widths = sum(d['box_w'] for d in vp_sorted)
n = len(vp_sorted)
gap = (total_span - sum_widths) / (n - 1) if n > 1 else 0.0

cursor = left_bound
for d in vp_sorted:
    d['new_u'] = cursor + d['box_w'] / 2.0
    cursor += d['box_w'] + gap

# ── 7. Identify largest viewport (by area) ───────────────────────────
largest = max(vp_data, key=lambda d: d['box_w'] * d['box_h'])

# ── 8. Apply in one transaction ──────────────────────────────────────
with revit.Transaction('Align Section Views'):

    # Move viewport boxes to new positions
    for d in vp_data:
        d['vp'].SetBoxCenter(XYZ(d['new_u'], d['new_v'], 0.0))

    doc.Regenerate()

    # Title alignment: 0.5" below the bottom of the largest viewport;
    # all other titles snap to the same V coordinate.
    largest_bottom_v = largest['new_v'] - largest['box_h'] / 2.0
    target_title_top = largest_bottom_v - HALF_INCH_FT

    for d in vp_data:
        vp = d['vp']
        try:
            lbl = vp.GetLabelOutline()
            lbl_h = lbl.MaximumPoint.Y - lbl.MinimumPoint.Y
            lbl_center_v = (lbl.MinimumPoint.Y + lbl.MaximumPoint.Y) / 2.0
            desired_v = target_title_top - lbl_h / 2.0
            delta_v = desired_v - lbl_center_v
            offset = vp.LabelOffset
            vp.LabelOffset = XYZ(offset.X, offset.Y + delta_v, 0.0)
        except Exception:
            pass  # LabelOffset unavailable or viewport has no label

# ── 9. Report ────────────────────────────────────────────────────────
output.print_md('**Done.** Aligned {} viewports.'.format(len(vp_data)))
output.print_md('Reference level: **{}** ({:.3f} ft)'.format(ref_level.Name, ref_level.Elevation))
