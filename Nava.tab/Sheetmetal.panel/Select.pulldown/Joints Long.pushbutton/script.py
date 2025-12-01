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
from revit_duct import JointSize, RevitDuct
from revit_output import print_parameter_help
from pyrevit import revit, script
from Autodesk.Revit.DB import *

# Button info
# ===================================================
__title__ = "Joints Long"
__doc__ = """
Selects joints that are MORE than:
TDF     = 56"
S&D     = 59"
Spiral  = 120"
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

# Filter ducts by size
fil_ducts = [d for d in ducts if d.joint_size == JointSize.LONG]

# Start of our logic / print
if fil_ducts:

    # Select filtered duct list
    RevitElement.select_many(uidoc, fil_ducts)
    output.print_md("# Selected {} long joints".format(len(fil_ducts)))
    output.print_md(
        "------------------------------------------------------------------------------")

    # loop for individutal duct and their selected properties
    for i, fil in enumerate(fil_ducts, start=1):
        output.print_md('### Index: {} | Lenght: {}" | Size: {} | Connectors: {}, {} | Element ID: {}'.format(
            i, fil.length, fil.size, fil.connector_0_type, fil.connector_1_type, output.linkify(
                fil.element.Id)
        ))

    # loop for totals
    element_ids = [d.element.Id for d in fil_ducts]
    output.print_md("# Total elements: {}, {}".format(
        len(fil_ducts), output.linkify(element_ids)
    ))

    # Final print statements
    print_parameter_help(output)
else:
    output.print_md("No short joints found")
