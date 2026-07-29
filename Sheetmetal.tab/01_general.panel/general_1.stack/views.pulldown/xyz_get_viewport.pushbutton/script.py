# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import revit, forms, script
from Autodesk.Revit.DB import Viewport, ViewSheet

# Button info
# ======================================================================
__title__ = 'Get Viewport XYZ'
__doc__ = '''
Print selected viewport center XYZ
'''

# Variables
# ======================================================================

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

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

for index, viewport in enumerate(viewports, 1):
    center = viewport.GetBoxCenter()
    output.print_md(
        '{}. Viewport {} center XYZ: ({:.16f}, {:.16f}, {:.16f})'.format(
            index,
            viewport.Id.IntegerValue,
            center.X,
            center.Y,
            center.Z,
        )
    )
