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
from revit_output import print_disclaimer
from pyrevit import revit, script
from Autodesk.Revit.DB import *

# Button info
# ===================================================
__title__ = "Unconnected"
__doc__ = """
Select all duct with an open connector
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
fil_ducts = [d for d in ducts if not d.fully_connected()]

# Start of select / print loop
if fil_ducts:

    # Select filtered dcuts
    RevitElement.select_many(uidoc, fil_ducts)
    output.print_md("# Selected {} unconnected duct".format(len(fil_ducts)))
    output.print_md("---")

    # Individutal duct and selected properties
    for i, fil in enumerate(fil_ducts, start=1):
        length_val = fil.length if fil.length else 0.0
        output.print_md(
            '### No: {:03} | ID: {} | Lenght: {:06.2f}" | Family: {} | Size: {}'.format(
                i,
                output.linkify(fil.element.Id),
                length_val,
                fil.family,
                fil.size,
            )
        )

    # Total count
    element_ids = [d.element.Id for d in fil_ducts]
    output.print_md(
        "# Total elements {}, {}".format(
            len(element_ids),
            output.linkify(element_ids)),
    )

    # Final print statements
    print_disclaimer(output)
else:
    output.print_md("## No unconnected ducts found")
