# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from revit.revit_element import RevitElement
from ducts.revit_duct import RevitDuct
from constants.print_outputs import print_disclaimer
from pyrevit import revit, script
from Autodesk.Revit.DB import *

# Button info
# ===================================================
__title__ = "Sleeve Duct"
__doc__ = """
Selects MEP ducts where `_type` is `sleeve`
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

# Get all duct
ducts = RevitDuct.all(doc, view)

# Filter ducts by `_type` parameter value
fil_ducts = [
    d for d in ducts
    if (d._get_param("_type") or "").strip().lower() == "sleeve"
]

# Start of select / print
if fil_ducts:

    # Select filtered duct
    RevitElement.select_many(uidoc, fil_ducts)
    output.print_md(
        "# Selected {} sleeve ducts".format(len(fil_ducts))
    )
    output.print_md("---")

    # Individual duct and properties
    for i, fil in enumerate(fil_ducts, start=1):
        output.print_md(
            '### No: {:03} | ID: {} | Length: {:06.2f}" | Size: {}'.format(
                i,
                output.linkify(fil.element.Id),
                fil.length,
                fil.size,
            )
        )

    element_ids = [d.element.Id for d in fil_ducts]
    output.print_md(
        "# Total elements {}, {}".format(
            len(fil_ducts),
            output.linkify(element_ids)
        )
    )

    # Final print statements
    print_disclaimer(output)
else:
    output.print_md("## No sleeve ducts found")
