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
Description: Test button. DO NOT PRESS.
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
    forms.alert("Please select one or more ducts first.")

family_list = ["transition", "mitred offset", "radius offset"]

for d in ducts:
    if d.family and d.family.strip().lower() in family_list:
        # Get connectors
        all_connectors = list(d.element.ConnectorManager.Connectors)
        c_0 = all_connectors[0] if len(all_connectors) > 0 else None
        c_1 = all_connectors[1] if len(all_connectors) > 1 else None

        # Sizes
        w_i = d.width_in
        h_i = d.heigth_in
        w_o = d.width_out
        h_o = d.heigth_out

        # Fallbacks
        if w_o is None:
            w_o = w_i
        if h_o is None:
            h_o = h_i

        if c_0 and c_1:
            # Revit internal units (feet) -> inches
            p0 = (c_0.Origin.X * 12.0, c_0.Origin.Y * 12.0, c_0.Origin.Z * 12.0)
            p1 = (c_1.Origin.X * 12.0, c_1.Origin.Y * 12.0, c_1.Origin.Z * 12.0)

            # Use connector CS for width/height axes when available
            try:
                cs = c_0.CoordinateSystem
                u_hat = (cs.BasisX.X, cs.BasisX.Y, cs.BasisX.Z)  # width axis
                v_hat = (cs.BasisY.X, cs.BasisY.Y, cs.BasisY.Z)  # height axis
            except Exception:
                u_hat = (1.0, 0.0, 0.0)
                v_hat = (0.0, 1.0, 0.0)

            # Centerline offsets in inches
            delta = (p1[0] - p0[0], p1[1] - p0[1], p1[2] - p0[2])
            width_offset = abs(RevitXYZ.dot(delta, u_hat))
            height_offset = abs(RevitXYZ.dot(delta, v_hat))

            # Edge offsets (whole inches + exact)
            offsets = RevitXYZ.edge_diffs_whole_in(
                p0, w_i, h_i, p1, w_o, h_o, u_hat, v_hat)

            # Outputd
            output.print_md("\n**Duct: {}**".format(d.element.Id))
            output.print_md("Size: {}x{} to {}x{}".format(w_i, h_i, w_o, h_o))
            output.print_md("\n**Centerline Offsets:**")
            output.print_md("Width: {:.2f}\" | Height: {:.2f}\"".format(
                width_offset, height_offset))

            if offsets:
                output.print_md("\n**Edge Offsets:**")
                output.print_md("Left: {}\" | Right: {}\" | Top: {}\" | Bottom: {}\"".format(
                    offsets["whole_in"]["left"],
                    offsets["whole_in"]["right"],
                    offsets["whole_in"]["top"],
                    offsets["whole_in"]["bottom"]
                ))
