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
__title__ = 'Set View Location'
__doc__ = '''
Set view port to hard coded location
'''

# Variables
# ======================================================================

output = script.get_output()
doc = revit.doc
uidoc = revit.uidoc

TARGET_CENTER = XYZ(
    -0.0064048997808098251,
    0.0030527024873521214,
    0.39626736111096306,
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

with revit.Transaction('Set Viewport Center Location'):
    for viewport in viewports:
        viewport.SetBoxCenter(TARGET_CENTER)

# output.print_md(
#     'Moved {} viewport(s) to center: ({:.16f}, {:.16f}, {:.16f})'.format(
#         len(viewports),
#         TARGET_CENTER.X,
#         TARGET_CENTER.Y,
#         TARGET_CENTER.Z,
#     )
# )
