# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from revit_element import RevitElement
from revit_output import print_parameter_help
from revit_duct import RevitDuct
from pyrevit import revit, script
from Autodesk.Revit.DB import *
from collections import Counter


__title__ = "Accessories"
__doc__ = """
Selects all end caps and taps
"""

# Imports
# ==================================================

# .NET Imports
# ==================================================


# Variables
# ==================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
view = revit.active_view
output = script.get_output()

# Main Code
# ==================================================

# Gather ducts in the view
ducts = RevitDuct.all(doc, view)

# List of acceptable families / list of what families we are after
allowed = {"conicaltap - wdamper", "boot tap - wdamper",
           "8inch long coupler wdamper", "end cap", "cap",
           "tdf end cap"}

# Loops through all ducts and filters out famies not in our focus_families list
normalized = [(d, (d.family or "").lower().strip()) for d in ducts]
fil_ducts = [d for d, fam in normalized if fam in allowed]

# Start of our select / print loop
if fil_ducts:
    # Select all fitered duct
    RevitElement.select_many(uidoc, fil_ducts)
    output.print_md(
        "# Selected {} accessories  ".format(len(fil_ducts)))
    output.print_md(
        "------------------------------------------------------------------------------")

    # Individual links
    for i, d in enumerate(fil_ducts, start=1):
        output.print_md(
            "### Index: {} | Family: {} | Size: {} | Element ID: {}".format(
                i, d.family, d.size, output.linkify(d.element.Id)
            )
        )

    # Counters
    counts = Counter((d.family or "").lower().strip() for d in fil_ducts)

    # Final prints
    output.print_md(
        "### Selected {} conical taps  ".format(
            counts.get("conicaltap - wdamper", 0)))
    output.print_md(
        "### Selected {} boot tap  ".format(
            counts.get("boot tap - wdamper", 0)))
    output.print_md(
        "### Selected {} long coupler  ".format(
            counts.get("8inch long coupler wdamper", 0)))
    output.print_md(
        "### Selected {} end caps".format(
            counts.get("end cap", 0) + counts.get("cap", 0) + counts.get("tdf end cap", 0)))

    # Final print statements
    print_parameter_help(output)
else:
    output.print_md(
        "No accessories found.")
