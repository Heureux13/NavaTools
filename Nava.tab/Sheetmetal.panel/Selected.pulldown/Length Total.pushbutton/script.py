# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from System.Collections.Generic import List
from Autodesk.Revit.ApplicationServices import Application
from Autodesk.Revit.UI import UIDocument
from Autodesk.Revit.DB import ElementId, Document
from pyrevit import revit, forms, script, DB
from revit_duct import RevitDuct

__title__ = "Length Total"
__doc__ = """************************************************************************
Description:

Returns total length for ducts(s) selected
************************************************************************"""

# Variables
# ==================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
view = revit.active_view
output = script.get_output()

# Main Code
# ==================================================
ducts = RevitDuct.from_selection(uidoc, doc, view)

if not ducts:
    forms.alert("Please select one or more ducts first.")
else:
    # keep both the ElementId and the length
    lengths = [(d.element.Id, d.length)
               for d in ducts if d.length is not None]

    # Section title
    output.print_md("### Total Lengths")

    # Individual links with lengths
    for eid, l in lengths:
        output.print_md("- {}: {:.3f} in".format(output.linkify(eid), l))

    # Select All link
    all_ids = List[ElementId]()
    for eid, _ in lengths:
        all_ids.Add(eid)
    output.print_md("**{}**".format(output.linkify(all_ids)))

    # Footer total
    total = sum(w for _, w in lengths)
    output.print_md('**➡️ Total Duct Length: {:.3f} ft**'.format(total/12))