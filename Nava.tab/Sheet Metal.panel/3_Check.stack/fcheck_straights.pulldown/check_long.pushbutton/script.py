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
from ducts.revit_duct import CONNECTOR_THRESHOLDS, DEFAULT_SHORT_THRESHOLD_IN, RevitDuct
from constants.print_outputs import print_disclaimer
from pyrevit import revit, script
from Autodesk.Revit.DB import *

# Button info
# ===================================================
__title__ = "Long"
__doc__ = """
Selects Straight duct parts that are longer than connector thresholds.
TDF     = 56"
S&D     = 59"
Default = 56"
"""

# Variables
# ==================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
view = revit.active_view
output = script.get_output()
LONG_TOL_IN = 0.01

# Main Code
# ==================================================

# Get all ducts in view
ducts = RevitDuct.all(doc, view)

# Filter ducts by family, connector pair, and connector threshold
fil_ducts = []
threshold_by_id = {}
for d in ducts:
    family = (d.family or "").strip()
    conn0 = (d.connector_0_type or "").strip()
    conn1 = (d.connector_1_type or "").strip()

    if family != "Straight":
        continue
    if not conn0 or not conn1:
        continue
    if conn0 != conn1:
        continue
    if d.length is None:
        continue

    threshold = CONNECTOR_THRESHOLDS.get(("Straight", conn0), DEFAULT_SHORT_THRESHOLD_IN)
    if d.length > (threshold + LONG_TOL_IN):
        fil_ducts.append(d)
        threshold_by_id[d.id] = threshold

# Start of our logic / print
if fil_ducts:

    # Select filtered duct list
    RevitElement.select_many(uidoc, fil_ducts)
    output.print_md("# Selected {} long joints".format(len(fil_ducts)))
    output.print_md(
        "---")

    # loop for individutal duct and their selected properties
    for i, fil in enumerate(fil_ducts, start=1):
        threshold = threshold_by_id.get(fil.id, DEFAULT_SHORT_THRESHOLD_IN)
        output.print_md(
            '### Index: {:03} | Element ID: {} | Length: {:07.3f}" | Threshold: {:06.2f}" | Size: {} | Connectors: {}, {}'.format(
                i,
                output.linkify(
                    fil.element.Id),
                fil.length,
                threshold,
                fil.size,
                fil.connector_0_type,
                fil.connector_1_type,
            ))

    # loop for totals
    element_ids = [d.element.Id for d in fil_ducts]
    output.print_md("# Total elements: {}, {}".format(
        len(fil_ducts), output.linkify(element_ids)
    ))

    # Final print statements
    print_disclaimer(output)
else:
    output.print_md("No straight joints longer than threshold found")
