# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import script, revit

# Button info
# ======================================================================
__title__ = 'Get Scope Box Size'
__doc__ = 'Prints the dimensions of selected scope boxes'

# Variables
# ======================================================================
doc = revit.doc  # type: ignore[attr-defined]
uidoc = revit.uidoc  # type: ignore[attr-defined]
output = script.get_output()

elements = revit.get_selection().elements

if not elements:
    output.print_md('No elements selected.')
else:
    for el in elements:
        bbox = el.get_BoundingBox(None)

        if bbox:
            min_pt = bbox.Min
            max_pt = bbox.Max

            width = max_pt.X - min_pt.X
            height = max_pt.Y - min_pt.Y
            depth = max_pt.Z - min_pt.Z

            output.print_md(
                '**Element: {}**'.format(el.Name if hasattr(el, 'Name') else el.Id))
            output.print_md('Width (X): {:.2f}'.format(width))
            output.print_md('Height (Y): {:.2f}'.format(height))
            output.print_md('Depth (Z): {:.2f}'.format(depth))
            output.print_md('---')
        else:
            output.print_md(
                'No bounding box for element: {}'.format(el.Id))
