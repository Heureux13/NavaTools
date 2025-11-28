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
from revit_duct import RevitDuct, CONNECTOR_THRESHOLDS
from Autodesk.Revit.ApplicationServices import Application
from Autodesk.Revit.UI import UIDocument
from pyrevit import revit, forms, script
from Autodesk.Revit.DB import *
from System.Collections.Generic import List

# Button info
# ===================================================
__title__ = "Joint Spiral"
__doc__ = """
************************************************************************
Description:

Selects all spiral joints.
************************************************************************
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

# Get all ducts
ducts = RevitDuct.all(doc, view)
allowed_joints = {
    ("tube", "grc_swage-female"),
    ("spiral duct", "raw"),
    ("spiral pipe", "raw")
}

# Normalize family and connector
normalized = [(d, (d.family or "").lower().strip(),
               (d.connector_0_type or "").lower().strip()) for d in ducts]

# Filter using normalized values
fil_ducts = [d for d, fam, conn in normalized if (fam, conn) in allowed_joints]

# Start of select / print logic
if fil_ducts:

    # slect filtered ducts
    RevitElement.select_many(uidoc, fil_ducts)
    output.print_md(
        "# Selected {} spiral joints".format(
            len(
                fil_ducts
            )
        )
    )
    output.print_md(
        "----------------------------------------------------"
    )

    # Individual duct and selected parameters
    for i, fil in enumerate(
        fil_ducts, start=1
    ):
        output.print_md(
            '### Index: {} | Size: {} | Length: {}" | Element ID: {}'.format(
                i, fil.size, fil.length, output.linkify(fil.element.Id)
            )
        )

    # Total count
    element_ids = [d.element.Id for d in fil_ducts]
    output.print_md(
        '# Total elements: {}, {}'.format(
            len(element_ids), output.linkify(element_ids)
        )
    )

    # Final print statements
    output.print_md(
        "------------------------------------------------------------------------------")
    output.print_md(
        "If info is missing, make sure you have the parameters turned on from Naviate")
    output.print_md(
        "All from Connectors and Fabrication, and size from Fab Properties")
else:
    output.print_md("## No spiral joints selected")
