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
__title__ = "Joints TDF"
__doc__ = """
Selects all TDF joints
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

# Families allowed
allowed = {
    ("straight", "tdc"),
    ("straight", "tdf"),
}

# Nomalize and filter duct
normalized = [(
    d, (d.family or "").lower().strip(
    ), (d.connector_0_type or "").lower().strip()
) for d in ducts]
fil_ducts = [d for d, fam, conn in normalized if (fam, conn) in allowed]

# Start of select / print
if fil_ducts:

    # Select filtered duct
    RevitElement.select_many(uidoc, fil_ducts)
    output.print_md(
        "# Selected {} TDF straight joints".format(len(fil_ducts))
    )
    output.print_md("---")

    # Individual duct and properties
    for i, fil in enumerate(fil_ducts, start=1):
        output.print_md(
            '### Index: {} | Size: {} | Length: {}" | Element ID: {}'.format(
                i, fil.size, fil.length, output.linkify(fil.element.Id)
            )
        )

    element_ids = [d.element.Id for d in fil_ducts]
    output.print_md(
        "# Total elements {}, {}".format(
            len(fil_ducts), output.linkify(element_ids)
        )
    )

    # Final print statements
    print_parameter_help(output)
else:
    output.print_md("## No TDF joints found")
