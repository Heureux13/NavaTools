# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================


from pyrevit import revit, script
from revit_xyz import RevitXYZ


# Button info
# ===================================================
__title__ = "XYZ"
__doc__ = """
Gets the XYZ Coordinates for Revit elements
"""

# Variables
# ==================================================
output = script.get_output()

# Main
# ==================================================
selection = revit.get_selection()

if not selection:
    output.print_md("Select at least one element.")
else:
    for el in selection:
        xyz = RevitXYZ(el)

        # Try curve endpoints first
        sp, ep = xyz.curve_endpoints()
        if sp and ep:
            output.print_md(
                "Element {}: start=({:.3f}, {:.3f}, {:.3f}), end=({:.3f}, {:.3f}, {:.3f})".format(
                    el.Id.Value,
                    sp.X, sp.Y, sp.Z,
                    ep.X, ep.Y, ep.Z,
                ))

        else:
            # Try connector origins
            origins = xyz.connector_origins()
            if origins:
                for i, o in enumerate(origins):
                    output.print_md(
                        "Element {}: connector {} origin=({:.3f}, {:.3f}, {:.3f})".format(
                            el.Id.Value,
                            i,
                            o.X,
                            o.Y,
                            o.Z
                        ))
            else:
                output.print_md(
                    "Element {}: no Location.Curve and no connectors".format(el.Id.Value))
