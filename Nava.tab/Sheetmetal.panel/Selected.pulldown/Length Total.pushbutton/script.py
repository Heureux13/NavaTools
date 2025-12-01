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
from revit_output import print_parameter_help
from revit_duct import RevitDuct

__title__ = "Length Total"
__doc__ = """
******************************************************************
Description:
Returns length for duct(s) selected.
******************************************************************
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
    output.print_md(
        '---------------------------------------------------------')

    # Individual properties
    for i, d in enumerate(ducts, start=1):
        output.print_md(
            '### Index: {} | Size: {} | Length: {}" | Family: {} | Element ID: {}'.format(
                i, d.size, d.length, d.family, output.linkify(d.element.Id)
            )
        )

    # Final totals loop and link
    element_ids = [d.element.Id for d in ducts]
    lengths = [d.length for d in ducts if d.length is not None]
    output.print_md(
        '# Total elements: {} | Total length: {:.2f} ft | {}'.format(
            len(element_ids), sum(lengths) / 12, output.linkify(element_ids)
        )
    )

    # Final print statements
    print_parameter_help(output)
else:
    output.print_md("No ductwork found.")
