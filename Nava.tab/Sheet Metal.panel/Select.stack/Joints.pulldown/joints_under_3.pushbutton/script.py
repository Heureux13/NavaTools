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
__title__ = 'Spiral Under 3" Long'
__doc__ = """
placeholder
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
duct = RevitDuct.from_selection(uidoc, doc, view)

# Filter down to short joints
DUCT_LENGTH = 6

normalize = [
    (d.family.strip().lower(), d.length) for d in ducts if d.family and d.length
]

acceptable = {
    "spiral duct",
    "straight",
}

fil_ducts = [
    d for d in ducts
    if d.length and d.length < DUCT_LENGTH and d.family and d.family.strip().lower() in acceptable
]

# Start of select / print loop
if fil_ducts:

    # Select filtered dcuts
    RevitElement.select_many(uidoc, fil_ducts)
    output.print_md(
        '# Selected {:03} sprial joints under 3" long'.format(len(fil_ducts)))
    output.print_md("---")

    # Individutal duct and selected properties
    for i, fil in enumerate(fil_ducts, start=1):
        output.print_md(
            '### No: {:03} | ID: {} | Size: {} | Length: {:06.2f}"'.format(
                i,
                output.linkify(fil.element.Id),
                fil.size,
                fil.length,
            )
        )

    # Total count
    element_ids = [d.element.Id for d in fil_ducts]
    output.print_md(
        "# Total elements {:03}, {}".format(
            len(element_ids),
            output.linkify(element_ids))
    )

    # Final print statements
    print_parameter_help(output)
else:
    output.print_md('## No straight duct under 6" found')
