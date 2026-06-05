# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import revit, forms, script
from Autodesk.Revit.DB import (
    Viewport,
    ViewSheet,
    XYZ,
)

# Button info
# ======================================================================
__title__ = 'Align Views Vertically'
__doc__ = '''
Select 2+ section viewports on an active sheet, then run.
  - Aligns viewports horizontally so their level lines share the same U
  - Equally spaces viewports vertically (outer bounds preserved)
  - Places every viewport title 0.5" below its view, left-aligned to the largest
'''

# Constants
# ======================================================================
HALF_INCH_FT = 0.15 / 12.0

# Variables
# ======================================================================
output = script.get_output()
doc = revit.doc
uidoc = revit.uidoc

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
        box_outline = vp.GetBoxOutline()
        box_center = vp.GetBoxCenter()
        box_w = box_outline.MaximumPoint.X - box_outline.MinimumPoint.X
        box_h = box_outline.MaximumPoint.Y - box_outline.MinimumPoint.Y

        vp_data.append({
            'vp': vp,
            'box_center': box_center,
            'box_w': box_w,
            'box_h': box_h,
        })
    except Exception as exc:
        output.print_md('Skipped viewport {}: {}'.format(vp.Id.IntegerValue, exc))

if len(vp_data) < 2:
    forms.alert(
        'Need at least 2 valid section viewports with crop boxes.',
        exitscript=True
    )

# ── 4. (no level reference needed for vertical layout) ──────────────

# ── 5. Keep horizontal positions unchanged ───────────────────────────
for d in vp_data:
    d['new_u'] = d['box_center'].X

# ── 6. Equal vertical spacing – compute new V centers ────────────────
# Sort top-to-bottom by current box center V (descending)
vp_sorted = sorted(vp_data, key=lambda d: d['box_center'].Y, reverse=True)

top_bound = vp_sorted[0]['box_center'].Y + vp_sorted[0]['box_h'] / 2.0
bottom_bound = vp_sorted[-1]['box_center'].Y - vp_sorted[-1]['box_h'] / 2.0
total_span = top_bound - bottom_bound
sum_heights = sum(d['box_h'] for d in vp_sorted)
n = len(vp_sorted)
gap = (total_span - sum_heights) / (n - 1) if n > 1 else 0.0

cursor = top_bound
for d in vp_sorted:
    d['new_v'] = cursor - d['box_h'] / 2.0
    cursor -= d['box_h'] + gap

# ── 7. Identify largest viewport (by area) ───────────────────────────
largest = max(vp_data, key=lambda d: d['box_w'] * d['box_h'])

# ── 8. Apply in one transaction ──────────────────────────────────────
with revit.Transaction('Align Section Views Vertically'):

    # Move viewport boxes to new positions
    for d in vp_data:
        d['vp'].SetBoxCenter(XYZ(d['new_u'], d['new_v'], 0.0))

    doc.Regenerate()

    # Title alignment: 0.5" below each view's bottom; left-align all
    # title labels to the left edge of the largest viewport's title.
    try:
        lbl_largest = largest['vp'].GetLabelOutline()
        largest_title_left_u = lbl_largest.MinimumPoint.X
    except Exception:
        lbl_largest = None
        largest_title_left_u = None

    for d in vp_data:
        vp = d['vp']
        try:
            lbl = vp.GetLabelOutline()
            lbl_h = lbl.MaximumPoint.Y - lbl.MinimumPoint.Y
            lbl_center_v = (lbl.MinimumPoint.Y + lbl.MaximumPoint.Y) / 2.0
            box_bottom_v = d['new_v'] - d['box_h'] / 2.0
            desired_v = box_bottom_v - HALF_INCH_FT - lbl_h / 2.0
            delta_v = desired_v - lbl_center_v

            delta_u = 0.0
            if largest_title_left_u is not None:
                current_left_u = lbl.MinimumPoint.X
                delta_u = largest_title_left_u - current_left_u

            offset = vp.LabelOffset
            vp.LabelOffset = XYZ(offset.X + delta_u, offset.Y + delta_v, 0.0)
        except Exception:
            pass  # LabelOffset unavailable or viewport has no label

# ── 9. Report ────────────────────────────────────────────────────────
output.print_md('**Done.** Aligned {} viewports vertically.'.format(len(vp_data)))
doc = revit.doc
uidoc = revit.uidoc

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
        local_y_center = (crop_box.Min.Y + crop_box.Max.Y) / 2.0

        # World Z (elevation) at the vertical center of the crop box
        local_mid = XYZ(
            (crop_box.Min.X + crop_box.Max.X) / 2.0,
            local_y_center,
            0.0
        )
        world_mid = transform.OfPoint(local_mid)
        elev_center = world_mid.Z

        box_outline = vp.GetBoxOutline()
        box_center = vp.GetBoxCenter()
        box_w = box_outline.MaximumPoint.X - box_outline.MinimumPoint.X
        box_h = box_outline.MaximumPoint.Y - box_outline.MinimumPoint.Y

        vp_data.append({
            'vp': vp,
            'scale': scale,
            'elev_center': elev_center,
            'box_center': box_center,
            'box_w': box_w,
            'box_h': box_h,
        })
    except Exception as exc:
        output.print_md('Skipped viewport {}: {}'.format(vp.Id.IntegerValue, exc))

if len(vp_data) < 2:
    forms.alert(
        'Need at least 2 valid section viewports with crop boxes.',
        exitscript=True
    )

# ── 9. Report ────────────────────────────────────────────────────────
output.print_md('**Done.** Aligned {} viewports vertically.'.format(len(vp_data)))
