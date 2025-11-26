# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from pyrevit import revit, script, forms
from revit_xyz import RevitXYZ
from revit_duct import RevitDuct
from revit_parameter import RevitParameter
from Autodesk.Revit.DB import Transaction

# Button display information
# =================================================
__title__ = "DO NOT PRESS"
__doc__ = """
******************************************************************
Gives offset information about specific duct fittings
******************************************************************
"""

# Variables
# ==================================================
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
view = revit.active_view
output = script.get_output()
all_ducts = RevitDuct.all(doc, view)
ducts = RevitDuct.from_selection(uidoc, doc, view)
rp = RevitParameter(doc, app)

# Code
# =========================================================
if not ducts:
    output.print_md("selecte duct")
else:
    output.print_md('### Connector XYZ')
    for d in ducts:
        coords = d.connector_origins()
        output.print_md("**Elenet ID:** '{}'".formatn(d.id))
        if not coords:
            output.print_md("- No connecotr found")
            continue
        for i, (x, y, z) in enumerate(coords):
            output.print_md("- Connector {}: x={:.2f}, y={:.2f}, z={:.2f}".format(i, x, y, z))
