# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import script, revit
from revit.revit_views import RevitViews

# Button info
# ======================================================================
__title__ = 'moves viewport'
__doc__ = '''
moves the prot of di theintihivnghitghsh
'''

# Variables
# ======================================================================
doc = revit.doc  # type: ignore[attr-defined]
uidoc = revit.uidoc  # type: ignore[attr-defined]
view = doc.ActiveView
output = script.get_output()

rvt_viewport = RevitViews(doc, view)

viewport = rvt_viewport.get_viewport_info(doc, view)
output.print_md('viewport_id: {}'.format(viewport[0]['viewport_id']))
output.print_md('view_name: {}'.format(viewport[0]['view_name']))
output.print_md('center: {}'.format(viewport[0]['center']))
output.print_md('width: {}'.format(viewport[0]['width']))
output.print_md('height: {}'.format(viewport[0]['height']))
