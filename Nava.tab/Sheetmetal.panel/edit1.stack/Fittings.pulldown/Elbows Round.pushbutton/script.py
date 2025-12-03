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
from revit_duct import RevitDuct
from revit_output import print_parameter_help
from pyrevit import revit, script
from Autodesk.Revit.DB import *

# Button info
# ===================================================
__title__ = "Elbows Round"
__doc__ = """
Select all round elbows
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

# List of acceptable familes / list of what familes we are after
allowed = {"gored elbow"}

# Loops through all duct and filters down to only wanted families
normalized = [(d, (d.family or "").lower().strip()) for d in ducts]
sel_ducts = [d for d, fam in normalized if fam in allowed]

# start of our select / print loop
if sel_ducts:

    # Selctes the filtered ducts
    RevitElement.select_many(uidoc, sel_ducts)
    output.print_md("# Selected {:03} gored elbows".format(len(sel_ducts)))
    output.print_md("---")

    # Loop for individual duct and their selected properties
    for i, sel in enumerate(sel_ducts, start=1):
        output.print_md("### Index: {:03} | Size: {} | Angle: {} | Element ID: {}".format(
            i, sel.size, sel.angle, output.linkify(sel.element.Id)))

    # Loop for totals
    element_ids = [d.element.Id for d in sel_ducts]
    output.print_md("# Total elements: {:03}, {}".format(
        len(element_ids), output.linkify(element_ids)))

    # Final print statements
    print_parameter_help(output)
else:
    output.print_md("No round elbows found.")
