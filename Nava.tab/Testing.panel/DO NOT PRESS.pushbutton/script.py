# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
import math
from revit_duct import RevitDuct
from revit_xyz import RevitXYZ
from pyrevit import revit, script, forms

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
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
view = revit.active_view
output = script.get_output()

# Code
# ==================================================
ducts = RevitDuct.from_selection(uidoc, doc, view)
if not ducts:
    forms.alert("Select fabrication ducts first.", exitscript=True)

family_list = ["transition", "mitred offset",
               "radius offset", "mitered offset",
               "ogee", "offset", ]

for d in ducts:
    if d.family in family_list:
        cen_w = d.centerline_width
        cen_h = d.centerline_height
        off_l = d.offset_left
        off_r = d.offset_right
        off_t = d.offset_top
        off_b = d.offset_bottom
        top_e = d.topo_edge_rise_in()
        bot_e = d.bottom_edge_rise_in()

        if d.family == "transition":
            if cen_h == 0:
                tag = "CL"
            elif d.higher_connector_indes() == 0:
                tag = "{}↓".format(off_b)
            else:
                tag = "{}↑".format(off_t)
