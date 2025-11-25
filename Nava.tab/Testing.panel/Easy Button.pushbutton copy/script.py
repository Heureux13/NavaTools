# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from pyrevit import revit, script
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
ducts = RevitDuct.all(doc, view)
# ducts = RevitDuct.from_selection(uidoc, doc)
rp = RevitParameter(doc, app)


# Code
# ==================================================
with Transaction(doc, "Offset Parameter") as t:
    t.Start()

    # What a run is
    allowed_duct = [
        "straight", "radius elbow", "transition",
        "offset", "Tee", "elbow", "gored elbow",
        "reducer"
    ]

    seen_ids = set()

    # Find a starting point
    for d in ducts:
        for connector_index in [0, 1, 2]:
            connected_elements = d.get_connected_elements(connector_index)
            for elem in connected_elements:
                if elem.Id.IntegerValue in seen_ids:
                    continue

                connected_duct = RevitDuct(doc, view, elem)

                if connected_duct.family and connected_duct.family.strip().lower() in allowed_duct:
                    seen_ids.add(elem.Id.IntegerValue)

    #

    t.Commit()

try:
    conns = list(self.element.ConnectorManager.Connectors)
    if len(conns) < 2:
        return (None, None)
    c0, c1 = conns[0], conns[1]

    # Try rectangular sizes (inches)
    def rect_wh(conn):
        try:
            return conn.Width * 12.0, conn.Height * 12.0
        except Exception:
            return None, None

    w0, h0 = rect_wh(c0)
    w1, h1 = rect_wh(c1)

    # Try round diameters (inches)
    def diameter(conn):
        try:
            return conn.Radius * 24.0  # 2 * radius * 12
        except Exception:
            return None
    d0 = diameter(c0)
    d1 = diameter(c1)

    # Rectangular case first
    if w0 and h0 and w1 and h1:
        a0 = w0 * h0
        a1 = w1 * h1
        if abs(a0 - a1) > 1e-6:
            return (c0, c1) if a0 >= a1 else (c1, c0)
        # Tie: fall back to element id for stability
        return (c0, c1) if c0.Owner.Id.IntegerValue <= c1.Owner.Id.IntegerValue else (c1, c0)

    # Round case
    if d0 and d1:
        if abs(d0 - d1) > 1e-6:
            return (c0, c1) if d0 >= d1 else (c1, c0)
        return (c0, c1) if c0.Owner.Id.IntegerValue <= c1.Owner.Id.IntegerValue else (c1, c0)

    # Mixed or missing size info: fallback to id ordering
    return (c0, c1) if c0.Owner.Id.IntegerValue <= c1.Owner.Id.IntegerValue else (c1, c0)

except Exception:
    return (None, None)
