# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import script, revit
from revit.revit_element import RevitElement
from revit.revit_annotations import RevitAnnotations


# Button info
# ======================================================================
__title__ = 'Hides Number Tags'
__doc__ = '''
Temporarily hides all tags who's family is '_Tag.DCT_NumberDuct'.
'''

# Variables
# ======================================================================
doc = revit.doc
uidoc = revit.uidoc
plan_view = doc.ActiveView
output = script.get_output()


trigger_families = (
    '_Tag.DCT_NumberDuct',
)

annotations = RevitAnnotations(doc, plan_view, None)

tags_to_hide = annotations.get_tags_by_family(
    doc,
    plan_view,
    family=trigger_families,
)

output.print_md('hide {}'.format(len(tags_to_hide)))

if not tags_to_hide:
    output.print_md('No tags to hide')
else:
    with revit.Transaction('Hide tags'):
        hide_count = RevitElement.hide_elements_temp(plan_view, tags_to_hide)
    output.print_md('Temp tags hide {}'.format(hide_count))
