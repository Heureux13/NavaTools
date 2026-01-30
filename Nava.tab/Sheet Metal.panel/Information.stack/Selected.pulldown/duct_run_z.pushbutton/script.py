# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from revit_duct import RevitDuct
from pyrevit import revit, script
from System.Collections.Generic import List
from Autodesk.Revit.DB import ElementId

# Button display information
# =================================================
__title__ = "Select Run (Same Height)"
__doc__ = """
Select all ducts in a run at the same height based on size and connectors.
"""

# Code
# ==================================================
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
view = revit.active_view
output = script.get_output()

# Get selected ducts
selected_ducts = RevitDuct.from_selection(uidoc, doc, view)

if not selected_ducts:
    output.print_md("**No ducts selected.** Please select at least one duct.")
else:
    run_elements = set()
    all_run_ducts = []

    for selected_duct in selected_ducts:
        # Get the run at same height
        run = RevitDuct.create_duct_run_same_height(selected_duct, doc, view)

        # Add all elements from the run to selection set
        for duct in run:
            run_elements.add(duct.element.Id)
            if duct not in all_run_ducts:
                all_run_ducts.append(duct)

    # Convert set to .NET List for Revit selection
    run_ids = List[ElementId](run_elements)
    uidoc.Selection.SetElementIds(run_ids)

    # Print detailed information for each duct
    output.print_md("# Run Details")
    output.print_md("---")

    for i, sel in enumerate(all_run_ducts, start=1):
        length_val = sel.length or 0.0
        weight_val = sel.weight or 0.0
        size_val = sel.size or "N/A"

        length_str = "{:06.2f}".format(length_val)
        weight_str = "{:06.2f}".format(weight_val)

        lbs_per_ft = (weight_val / (length_val / 12.0)) if length_val else 0.0
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
    element_ids = [d.element.Id for d in all_run_ducts]
    total_length = sum(d.length or 0 for d in all_run_ducts)
    total_weight = sum(d.weight or 0 for d in all_run_ducts)
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
