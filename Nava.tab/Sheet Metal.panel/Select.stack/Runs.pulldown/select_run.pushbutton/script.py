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
__title__ = "Select Run"
__doc__ = """
Select joints based on size,
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
selected_duct = RevitDuct.from_selection(uidoc, doc, view)
selected_duct = selected_duct[0] if selected_duct else None

# Start of select / print loop
if selected_duct:
    # Selets duct that is connected to the selected duct based on size
    run = RevitDuct.create_duct_run(selected_duct, doc, view)
    RevitElement.select_many(uidoc, run)

    for i, sel in enumerate(run, start=1):
        output.print_md(
            '### No: {:03} | ID: {} | Size: {} | Length: {} | Weight {}'.format(
                i,
                output.linkify(sel.element.Id),
                sel.size,
                sel.length,
                sel.weight,
            )
        )

    # Total count
    element_ids = [d.element.Id for d in run]
    total_length = sum(d.length or 0 for d in run)
    total_weight = sum(d.weight or 0 for d in run)
    output.print_md(
        "# Total elements {:03} | Total Length: {} | Total Weight: {} | {}".format(
            len(element_ids),
            total_length,
            total_weight,
            output.linkify(element_ids)),
    )

    # Final print statements
    print_parameter_help(output)
else:
    output.print_md("## Select a duct first")
