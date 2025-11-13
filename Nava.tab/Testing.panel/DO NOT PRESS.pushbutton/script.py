# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================

from System.Collections.Generic import List
from revit_parameter import RevitParameter
from revit_element import RevitElement
from tag_duct import TagDuct
from revit_duct import RevitDuct, JointSize, CONNECTOR_THRESHOLDS
from revit_tagging import RevitXYZ
from Autodesk.Revit.ApplicationServices import Application
from Autodesk.Revit.UI import UIDocument
from pyrevit import revit, forms, script, DB
from Autodesk.Revit.DB import *
import clr

# Button display information
# =================================================
__title__ = "DO NOT PRESS"
__doc__ = """******************************************************************
Description:

Current goal fucntion of button is: select only spiral duct.

++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
This button is for testin code snippets and ideas before implementing them.
Odds are it will be constantly changing and not useful, its entire purpose
is for the author to have a quick button to test whatever code they are working on.
If you press it could do nothing, throw an error, or change something in your model.
Once working, it will most likely be moved to a more permanent location.
******************************************************************"""

# Variables
# ==================================================
app = __revit__.Application             # type: Application
uidoc = __revit__.ActiveUIDocument        # type: UIDocument
doc = revit.doc                         # type: Document
view = revit.active_view
output = script.get_output()

# Main Code
# ==================================================

# Get selected duct(s)
ducts = RevitDuct.from_selection(uidoc, doc)

if not ducts:
    forms.alert("please select one or more duct elements", exitscript=True)

# Header of pop up message
output.print_md("# XYZ of selected ducts")
output.print_md(
    "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++\n")

for duct in ducts:
    xyz = RevitXYZ(duct.element)

    start = xyz.start_point()
    mid = xyz.mid_point()
    end = xyz.end_point()

    # Print duct info
    output.print_md("## Duct ID: {}".format(duct.id))

    if start:
        output.print_md("- Start Point: X: {:.2f}, Y: {:.2f}, Z: {:.2f}".format(
            start.X, start.Y, start.Z
        ))
    if mid:
        output.print_md("- Mid Point: X: {:.2f}, Y: {:.2f}, Z: {:.2f}".format(
            mid.X, mid.Y, mid.Z
        ))
    if end:
        output.print_md("- End Point: X: {:.2f}, Y: {:.2f}, Z: {:.2f}".format(
            end.X, end.Y, end.Z
        ))
    diff = start.X - end.X if start and end else None
    output.print_md("Difference is {:.3f} feet".format(diff))
    output.print_md("\n---\n")

output.print_md("**Total duct elements processed:** {}".format(len(ducts)))
