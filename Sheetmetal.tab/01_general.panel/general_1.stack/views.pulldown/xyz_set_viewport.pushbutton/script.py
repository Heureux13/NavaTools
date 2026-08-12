# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import revit, forms, script
from Autodesk.Revit.DB import Viewport, ViewSheet, XYZ

# Button info
# ======================================================================
__title__ = 'Set XYZ Viewport'
__doc__ = '''
Set view port to hard coded location
'''

# Variables
# ======================================================================

output = script.get_output()
doc = revit.doc
uidoc = revit.uidoc

# Target top-left corner of the viewport rectangle on the sheet.
TARGET_TOP_LEFT = XYZ(
    -1.568398,
    1.205570,
    0,
)


active_view = revit.active_view
if not isinstance(active_view, ViewSheet):
    forms.alert(
        'Active view must be a sheet.\nOpen a sheet and try again.',
        exitscript=True,
    )

sel_ids = uidoc.Selection.GetElementIds()
viewports = []
for eid in sel_ids:
    element = doc.GetElement(eid)
    if isinstance(element, Viewport):
        viewports.append(element)

if not viewports:
    forms.alert('Select at least one viewport on the sheet.', exitscript=True)

with revit.Transaction('Set Viewport Top-Left Corner'):
    for viewport in viewports:
        outline = viewport.GetBoxOutline()
        min_pt = outline.MinimumPoint
        max_pt = outline.MaximumPoint
        width = max_pt.X - min_pt.X
        height = max_pt.Y - min_pt.Y

        target_center = XYZ(
            TARGET_TOP_LEFT.X + (width / 2.0),
            TARGET_TOP_LEFT.Y - (height / 2.0),
            0,
        )

        viewport.SetBoxCenter(target_center)

# output.print_md(
#     'Moved {} viewport(s) to top-left corner: ({:.6f}, {:.6f}, {:.6f})'.format(
#         len(viewports),
#         TARGET_TOP_LEFT.X,
#         TARGET_TOP_LEFT.Y,
#         TARGET_TOP_LEFT.Z,
#     )
# )
