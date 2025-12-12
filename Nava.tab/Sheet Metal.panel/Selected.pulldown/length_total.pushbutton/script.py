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
        '---')

    # Helpers
    def _to_float(val):
        try:
            return float(val)
        except Exception:
            return None

    def _best_length_in(d):
        """Prefer centerline length; fall back to length (inches)."""
        for cand in (getattr(d, 'centerline_length', None), getattr(d, 'length', None)):
            num = _to_float(cand)
            if num is not None:
                return num
        return None

    # Individual properties
    for i, d in enumerate(ducts, start=1):
        num_len = _best_length_in(d)
        if num_len is not None:
            length_text = '{:06.2f}"'.format(num_len)
        else:
            # show raw param if present, otherwise N/A
            raw = getattr(d, 'centerline_length', None)
            if raw is None:
                raw = getattr(d, 'length', None)
            length_text = str(raw) if raw is not None else 'N/A'
        output.print_md(
            '### Index: {:04d} | Length: {} | Size: {} | Family: {} | Element ID: {}'.format(
                i,
                length_text,
                d.size,
                d.family,
                output.linkify(d.element.Id)
            )
        )

    # Final totals loop and link
    element_ids = [d.element.Id for d in ducts]
    lengths = [fl for fl in (_best_length_in(d) for d in ducts) if fl is not None]
    output.print_md(
        '# No: {} | ID: {} | Total length: {:.2f} ft'.format(
            len(element_ids),
            output.linkify(element_ids),
            (sum(lengths) / 12.0) if lengths else 0.0
        )
    )

    # Final print statements
    print_disclaimer(output)
else:
    output.print_md("No ductwork found.")
