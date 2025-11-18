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

# Helper Functions
# ==================================================


def normalize(v):
    """Normalize a 3D vector to unit length."""
    x, y, z = v
    m = math.sqrt(x*x + y*y + z*z)
    if m == 0:
        return (0, 0, 0)
    return (x/m, y/m, z/m)


def dot(a, b):
    """Dot product of two 3D vectors."""
    return a[0]*b[0] + a[1]*b[1] + a[2]*b[2]


def edge_diffs_whole_in(c0, W0, H0, c1, W1, H1, u_hat, v_hat, snap_tol=1/16.0):
    """
    Calculate edge-to-edge offsets for duct transitions.

    Parameters:
        c0, c1: Connector center positions (x, y, z) in inches
        W0, H0, W1, H1: Widths and heights in inches
        u_hat: Width axis unit vector (parallel to view on plan)
        v_hat: Height axis unit vector (perpendicular to u_hat on plan)
        snap_tol: Values below this snap to 0 before rounding (default 1/16")

    Returns:
        Dictionary with 'exact_in' and 'whole_in' offsets for left/right/top/bottom edges
    """
    # Check for None values
    if None in [W0, H0, W1, H1]:
        return None

    u_hat = normalize(u_hat)
    v_hat = normalize(v_hat)

    # Scalar projections of centers along u and v
    u0 = dot(c0, u_hat)
    u1 = dot(c1, u_hat)
    v0 = dot(c0, v_hat)
    v1 = dot(c1, v_hat)

    exact = {
        "left":   abs((u1 - W1/2.0) - (u0 - W0/2.0)),
        "right":  abs((u1 + W1/2.0) - (u0 + W0/2.0)),
        "top":    abs((v1 + H1/2.0) - (v0 + H0/2.0)),
        "bottom": abs((v1 - H1/2.0) - (v0 - H0/2.0)),
    }

    def snap_round(v):
        v2 = 0.0 if v < snap_tol else v
        return int(round(v2))

    whole = {k: snap_round(v) for k, v in exact.items()}
    return {"exact_in": exact, "whole_in": whole}


# Code
# ==================================================
ducts = RevitDuct.from_selection(uidoc, doc, view)

if not ducts:
    forms.alert("Please select one or more ducts first.")

family_list = ["transition", "mitred offset", "radius offset"]

for d in ducts:
    if d.family and d.family.strip().lower() in family_list:

        # Get all connectors as a list
        all_connectors = list(d.element.ConnectorManager.Connectors)
        c_0 = all_connectors[0] if len(all_connectors) > 0 else None
        c_1 = all_connectors[1] if len(all_connectors) > 1 else None

        w_i = d.width_in
        h_i = d.heigth_in
        w_o = d.width_out
        h_o = d.heigth_out

        # Fallback: if outlet dimensions are None, use inlet dimensions
        if w_o is None:
            w_o = w_i
        if h_o is None:
            h_o = h_i

        if c_0 and c_1:
            # Get connector origins - Revit coordinates are in FEET, convert to INCHES
            p0 = (c_0.Origin.X * 12.0, c_0.Origin.Y * 12.0, c_0.Origin.Z * 12.0)
            p1 = (c_1.Origin.X * 12.0, c_1.Origin.Y * 12.0, c_1.Origin.Z * 12.0)

            # Get the duct's coordinate system from the first connector
            try:
                cs = c_0.CoordinateSystem
                u_hat = (cs.BasisX.X, cs.BasisX.Y, cs.BasisX.Z)  # width axis
                v_hat = (cs.BasisY.X, cs.BasisY.Y, cs.BasisY.Z)  # height axis
            except Exception:
                # Fallback to plan view axes
                u_hat = (1.0, 0.0, 0.0)
                v_hat = (0.0, 1.0, 0.0)

            # Calculate centerline offset (perpendicular to duct axis)
            delta = (p1[0] - p0[0], p1[1] - p0[1], p1[2] - p0[2])
            width_offset = abs(dot(delta, u_hat))
            height_offset = abs(dot(delta, v_hat))

            # Calculate edge offsets
            offsets = edge_diffs_whole_in(
                p0, w_i, h_i, p1, w_o, h_o, u_hat, v_hat)

            # Print results
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
