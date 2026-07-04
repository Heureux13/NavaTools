# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import revit, script
from System.Collections.Generic import List
from revit.revit_element import RevitElement
from revit.revit_views import RevitViews


# Button info
# ======================================================================
__title__ = 'Hide Section View Markers'
__doc__ = '''
Hides section view markers not in triggers
'''

# Variables
# ======================================================================

doc = revit.doc  # type: ignore[attr-defined]
uidoc = revit.uidoc  # type: ignore[attr-defined]
plan_view = doc.ActiveView
output = script.get_output()

trigger_views = (
    'horizontal',
    'vertical',
)

trigger_keywords = (
    'skip',
)

section_views = RevitViews(doc, plan_view)

views_to_hide = section_views.get_views_in_view(
    doc,
    plan_view,
    key_name=trigger_views,
    keywords=trigger_keywords,
)

# output.print_md('hide {}'.format(len(views_to_hide)))

if not views_to_hide:
    output.print_md('No views to hide')
else:
    with revit.Transaction('temp hide test'):
        hide_count = RevitElement.hide_elements_temp(plan_view, views_to_hide)
    # output.print_md('Temp views hide {}'.format(hide_count))
