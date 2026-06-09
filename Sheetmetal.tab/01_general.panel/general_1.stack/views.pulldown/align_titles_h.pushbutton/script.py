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
__title__ = 'Align Titles Horizontally'
__doc__ = '''
Select 2+ viewports on an active sheet, then run.
  - Aligns viewport titles to one horizontal line (same Y)
  - Does not move section views
'''

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

# ── 3. Gather title data ─────────────────────────────────────────────
title_data = []
for vp in viewports:
    try:
        lbl = vp.GetLabelOutline()
        title_top = lbl.MaximumPoint.Y
        lbl_center_v = (lbl.MinimumPoint.Y + lbl.MaximumPoint.Y) / 2.0
        box_center_x = vp.GetBoxCenter().X
        title_data.append({
            'vp': vp,
            'title_top': title_top,
            'lbl_center_v': lbl_center_v,
            'box_center_x': box_center_x,
        })
    except Exception:
        # Ignore viewports without a usable title outline
        pass

if len(title_data) < 2:
    forms.alert(
        'Need at least 2 selected viewports with visible titles.',
        exitscript=True
    )

# Keep the left-most title as the alignment anchor so one title stays put.
anchor = min(title_data, key=lambda d: d['box_center_x'])
target_title_top = anchor['title_top']

# ── 4. Apply title-only alignment ────────────────────────────────────
with revit.Transaction('Align Viewport Titles Horizontally'):
    for d in title_data:
        vp = d['vp']
        delta_v = target_title_top - d['title_top']
        offset = vp.LabelOffset
        vp.LabelOffset = XYZ(offset.X, offset.Y + delta_v, offset.Z)

# ── 5. Report ────────────────────────────────────────────────────────
output.print_md('**Done.** Aligned {} viewport titles.'.format(len(title_data)))
