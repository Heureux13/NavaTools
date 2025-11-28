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
from pyrevit import revit, script
from Autodesk.Revit.DB import *

# Button info
# ===================================================
__title__ = "Joints S&D"
__doc__ = """
******************************************************************
Description:

Selects all S&D joints
******************************************************************
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

# Get all ducts in view
ducts = RevitDuct.all(doc, view)

# Family / connector combo to find
allowed = {
    ("straight", "slip & drive"),
    ("straight", "s&d"),
    ("straight", "standing s&d")
}

# List of filtered ducts
normalized = [(d, (d.family or "").lower().strip(), (d.connector_0_type or "").lower().strip()) for d in ducts]
fil_ducts = [d for d, fam, conn in normalized if (fam, conn) in allowed]

# Start of select / print loop
if fil_ducts:

    # Select filtered ducs
    RevitElement.select_many(uidoc, fil_ducts)
    output.print_md("# Selected {} S&D straight joints".format(len(fil_ducts)))
    output.print_md("------------------------------------------------------------------------------")

    # Loop for individutal duct and their selected properties
    for i, fil in enumerate(fil_ducts, start=1):
        output.print_md('### Index: {} | Size: {} | Element ID: {}'.format(
            i, fil.size, output.linkify(fil.element.Id)
            ))

    # Loop for total count
    element_ids = [d.element.Id for d in fil_ducts]
    output.print_md("# Total elements: {}, {}".format(
        len(element_ids), output.linkify(element_ids)
    ))

     # Final print statements
    output.print_md("------------------------------------------------------------------------------")
    output.print_md("If info is missing, make sure you have the parameters turned on from Naviate")
    output.print_md("All from Connectors and Fabrication, and size from Fab Properties")
else:
    output.print_md("## No S&D joints found")