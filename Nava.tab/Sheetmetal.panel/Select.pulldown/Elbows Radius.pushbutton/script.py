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
from Autodesk.Revit.ApplicationServices import Application
from Autodesk.Revit.UI import UIDocument
from pyrevit import revit, script, DB
from Autodesk.Revit.DB import *

# Button info
# ===================================================
__title__ = "Elbows Radius"
__doc__ = """
************************************************************************
Description:

Select all radius elbows
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
# Gathers Duct in the view
ducts = RevitDuct.all(doc, view)

# List of acceptable families / list of what familes we are after
allowed = {"radius elbow"}

# Normalize and filter
normalized = [(d, (d.family or "").lower().strip()) for d in ducts]
fil_ducts = [d for d, fam in normalized if fam in allowed]

# start of our select / print loop
if fil_ducts:

    # Select filtered ducts
    RevitElement.select_many(uidoc, fil_ducts)
    output.print_md("# Selected {} radius elbows".format(len(fil_ducts)))
    output.print_md(
        "------------------------------------------------------------------------------")

    # Loop for individudal duct and their selected properties
    for i, sel in enumerate(fil_ducts, start=1):
        output.print_md(
            "### Index: {} | Size: {} | Angle: {} | Inner Radius: {} | Element ID: {}".format(
                i, sel.size, sel.angle, sel.inner_radius, output.linkify(sel.element.Id)))

    # Loop for total counts
    element_ids = [d.element.Id for d in fil_ducts]
    output.print_md(
        "# Total elements: {}, {}".format(
            len(element_ids), output.linkify(element_ids)))

    # Final prints
    output.print_md(
        "------------------------------------------------------------------------------")
    output.print_md(
        "If info is missing, make sure you have the parameters turned on from Naviate")
    output.print_md(
        "All from Connectors and Fabrication, and size from Fab Properties")
else:
    output.print_md("No radius elbows found.")
