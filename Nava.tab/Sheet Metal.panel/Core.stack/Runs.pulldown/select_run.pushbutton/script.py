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
__title__ = "Select Run"
__doc__ = """
Selects/creates a run bases on size of seleted duct.
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
        length_val = sel.length if sel.length else 0.0
        weight_val = sel.weight if sel.weight else 0.0
        lbs_per_ft = (weight_val / (length_val / 12.0)) if length_val else 0.0
        size_val = sel.size if sel.size else "Unknown"

        length_str = "{:06.2f}".format(float(length_val))
        weight_str = "{:06.2f}".format(float(weight_val))
        lbs_ft_str = "{:06.2f}".format(float(lbs_per_ft))
        output.print_md(
            '### No: {:03} | ID: {} | Length: {} | Weight {} | lbs/ft: {} | Size: {}'.format(
                i,
                output.linkify(sel.element.Id),
                length_str,
                weight_str,
                lbs_ft_str,
                size_val,
            ))

    # Total count
    element_ids = [d.element.Id for d in run]
    total_length = sum(d.length or 0 for d in run)
    total_weight = sum(d.weight or 0 for d in run)
    total_lbs_per_ft = (total_weight / (total_length / 12.0)
                        ) if total_length else 0.0
    output.print_md("---")
    output.print_md(
        "# Total elements {} | Total feet: {} | Total lbs: {} | lbs/ft {} | {}".format(
            len(element_ids),
            "{:.2f}".format(total_length / 12.0),
            "{:.2f}".format(total_weight),
            "{:.2f}".format(total_lbs_per_ft),
            output.linkify(element_ids)),
    )

    # Final print statements
    print_disclaimer(output)
else:
    output.print_md("## Select a duct first")
