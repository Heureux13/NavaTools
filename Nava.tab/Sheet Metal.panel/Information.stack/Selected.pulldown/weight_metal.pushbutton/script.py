# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from pyrevit import revit, script
from revit_output import print_disclaimer
from revit_duct import RevitDuct

__title__ = "Total Weight"
__doc__ = """
Returns weight for duct(s) selected.
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
ducts = RevitDuct.from_selection(uidoc, doc, view)

# Select / print loop
if ducts:
    output.print_md('# Selected {} joints of duct'.format(len(ducts)))
    output.print_md('---')

    # Individual properties
    for i, d in enumerate(ducts, start=1):
        if len(ducts) < 501:
            output.print_md(
                '### No: {:03} | ID: {} | Weight: {:06.2f}lbs | Length: {:06.2f}" | Size: {} | Family: {}'.format(
                    i,
                    output.linkify(d.element.Id),
                    d.weight,
                    d.length,
                    d.size,
                    d.family
                )
            )

    # Final totals loop and link
    element_ids = [d.element.Id for d in ducts]
    pairs = [(d.weight, d.length)
             for d in ducts if d.weight is not None and d.length is not None]
    total_weight = sum(w for w, _ in pairs)
    total_length_in = sum(l for _, l in pairs)
    weight_per_ft = (total_weight / (total_length_in / 12.0)
                     ) if total_length_in else 0.0
    output.print_md(
        '# Total elements: {} | IDs: {} | Total weight: {:.2f} lbs | Weight/ft: {:.2f}'.format(
            len(element_ids),
            output.linkify(element_ids),
            total_weight,
            weight_per_ft
        )
    )

    # Final print statements
    print_disclaimer(output)
else:
    output.print_md("No ductwork found.")
