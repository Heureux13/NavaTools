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

__title__ = "Length Total"
__doc__ = """
Returns length for duct(s) selected.
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
    output.print_md('# Selected {} duct parts'.format(len(ducts)))
    output.print_md(
        '---')

    # Individual properties
    for i, d in enumerate(ducts, start=1):
        if len(ducts) < 501:
            output.print_md(
                '### Index: {:03d} | Element ID: {} | Length: {:06.2f}" | Size: {} | Family: {}'.format(
                    i,
                    output.linkify(d.element.Id),
                    d.length / 12 if d.length is not None else 0.00,
                    d.size,
                    d.family,
                )
            )

    # Final totals loop and link
    element_ids = [d.element.Id for d in ducts]
    total_ft = (sum(d.length for d in ducts if d.length is not None) / 12.0)
    total_ct = len(ducts)
    output.print_md(
        '# Total: {} | ID: {} | Total: {:.2f}ft | Average: {:06.2f}in'.format(
            len(element_ids),
            output.linkify(element_ids),
            total_ft,
            total_ft * 12 / total_ct if total_ct > 0 else 0,
        )
    )

    # Final print statements
    print_disclaimer(output)
else:
    output.print_md("No ductwork found.")
