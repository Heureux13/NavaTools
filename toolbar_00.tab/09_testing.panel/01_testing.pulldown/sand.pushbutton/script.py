# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.DB import BuiltInCategory
from pyrevit import script, revit
from ducts.revit_duct import RevitDuct

# Button info
# ======================================================================
__title__ = 'sand'
__author__ = 'Jose Francisco Nava Perez'
__doc__ = """
Prints the Element Id of each selected fabrication ductwork element."""

# Variables
# ======================================================================
uidoc = revit.uidoc
doc = revit.doc
output = script.get_output()

selection = revit.get_selection()

fab_ducts = [
    el for el in selection
    if el.Category
    and el.Category.Id.IntegerValue == int(BuiltInCategory.OST_FabricationDuctwork)
]

if not fab_ducts:
    output.print_md("**No fabrication ductwork selected.**")
    script.exit()

output.print_md("**Selected Fabrication Ductwork Element Ids:**")

for el in fab_ducts:
    d = RevitDuct(doc, revit.active_view, el)
    for p in el.Parameters:
        pname = p.Definition.Name
        output.print_md('{}: {}'.format(pname, d._get_param_v2(pname)))