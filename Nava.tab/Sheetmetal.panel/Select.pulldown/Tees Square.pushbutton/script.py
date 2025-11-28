# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
# Imports
# ==================================================
from revit_element import RevitElement
from revit_duct import RevitDuct
from Autodesk.Revit.ApplicationServices import Application
from Autodesk.Revit.UI import UIDocument
from pyrevit import revit, script
from Autodesk.Revit.DB import *


# Button info
# # ======================================================================
__title__ = "Tee Square"
__doc__ = """
****************************************************
Description:

Selects all square tees
****************************************************
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
allowed = {"tee"}

# Loops through all duct and filters down to only wanted families
normalized = [(d, (d.family or '').lower().strip()) for d in ducts]
sel_ducts = [d for d, fam in normalized if fam in allowed]

# Start of our select / print loop
if sel_ducts:

    # Selectes filtered ducts
    RevitElement.select_many(uidoc, sel_ducts)
    output.print_md("# Selected {} square tees".format(len(sel_ducts)))
    output.print_md(
        "------------------------------------------------------------------------------")

    # Loop for individual duct and their selected properties
    for i, sel in enumerate(sel_ducts, start=1):
        output.print_md(
            "### Index: {} | Size: {} | Throats: {}\", {}\", {}\" | Connectors: {}, {}, {} | Element ID: {}".format(
                i, sel.size, sel.extension_bottom, sel.extension_left,
                sel.extension_right, sel.connector_0_type, sel.connector_1_type,
                sel.connector_2_type, output.linkify(sel.element.Id)
            )
        )

    # Loop for total count
    element_ids = [d.element.Id for d in sel_ducts]
    output.print_md("# Total elements: {}, {}".format(
        len(element_ids), output.linkify(element_ids)
    ))

    # Final print statements
    output.print_md(
        "------------------------------------------------------------------------------")
    output.print_md(
        "If info is missing, make sure you have the parameters turned on from Naviate")
    output.print_md(
        "All from Connectors and Fabrication, and size from Fab Properties")
else:
    output.print_md("No square tees found.")
