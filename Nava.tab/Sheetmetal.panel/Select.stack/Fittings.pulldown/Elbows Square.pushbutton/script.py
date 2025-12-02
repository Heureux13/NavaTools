# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from revit_element import RevitElement
from revit_duct import RevitDuct, script
from revit_output import print_parameter_help
from pyrevit import revit
from Autodesk.Revit.DB import *


# Button info
# # ======================================================================
__title__ = "Elbows Square"
__doc__ = """
Selects all square elbows
"""

# Variables
# ==================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
view = revit.active_view
output = script.get_output()

# Main Code
# ==================================================

# Gathers Duct in the view
ducts = RevitDuct.all(doc, view)

# List of acceptable families / list of what familes we are after
allowed = {"elbow"}

# Loops through all duct and filters down to only wanted families
normalized = [(d, (d.family or '').lower().strip()) for d in ducts]
sel_ducts = [d for d, fam in normalized if fam in allowed]

# start of our select / print loop
if sel_ducts:

    # Selectes the filtered ducts
    RevitElement.select_many(uidoc, sel_ducts)
    output.print_md("# Selected {:03} square elbows".format(len(sel_ducts)))
    output.print_md(
        "---")

    # loop for individual duct and their selected properties
    for i, sel in enumerate(sel_ducts, start=1):
        output.print_md(
            "### Index: {:03} | Size: {} | Angle: {} | Type: {}, {} | Throat: {}, {} | Element ID: {}".format(
                i,
                sel.size,
                sel.angle,
                sel.connector_0_type,
                sel.connector_1_type,
                sel.extension_bottom,
                sel.extension_top,
                output.linkify(sel.element.Id)
            )
        )

    # Loop for totals
    element_ids = [d.element.Id for d in sel_ducts]
    output.print_md("# Total elements: {:03}, {}".format(
        len(element_ids), output.linkify(element_ids)))

    # Final print statements
    print_parameter_help(output)
else:
    output.print_md("No mitered elbows found.")
