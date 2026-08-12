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
__title__ = 'Get Viewport XYZ'
__doc__ = '''
Print selected viewport center, size, and corner coordinates.
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
    outline = None
    try:
        outline = viewport.GetBoxOutline()
    except Exception:
        outline = None

    width = None
    height = None
    corners = []

    if outline is not None:
        min_pt = outline.MinimumPoint
        max_pt = outline.MaximumPoint
        width = max_pt.X - min_pt.X
        height = max_pt.Y - min_pt.Y
        corners = [
            XYZ(min_pt.X, min_pt.Y, 0),
            XYZ(max_pt.X, min_pt.Y, 0),
            XYZ(max_pt.X, max_pt.Y, 0),
            XYZ(min_pt.X, max_pt.Y, 0),
        ]

    output.print_md(
        '{}. Viewport {} center XYZ: ({:.16f}, {:.16f}, {:.16f})'.format(
            index,
            viewport.Id.IntegerValue,
            center.X,
            center.Y,
            center.Z,
        )
    )

    if width is not None and height is not None:
        output.print_md(
            '   Size: width = {:.6f}, height = {:.6f} (sheet-space)'.format(width, height)
        )
        for corner_index, corner in enumerate(corners, 1):
            output.print_md(
                '   Corner {}: ({:.6f}, {:.6f}, {:.6f})'.format(
                    corner_index,
                    corner.X,
                    corner.Y,
                    corner.Z,
                )
            )
    else:
        output.print_md('   Size and corners could not be computed from the viewport outline.')
