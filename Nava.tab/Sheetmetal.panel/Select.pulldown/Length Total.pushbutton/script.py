# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

__title__   = "Length Total"
__doc__     = """
************************************************************************
Description:

Returns total length for ducts(s) selected
************************************************************************
"""

# Imports
# ==================================================
from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script, DB
from Autodesk.Revit.UI import UIDocument
from Autodesk.Revit.ApplicationServices import Application
from revit_duct import RevitDuct, JointSize
from tag_duct import TagDuct
from revit_element import RevitElement

#.NET Imports
# ==================================================
from System.Collections.Generic import List
import clr


# Variables
# ==================================================
app   = __revit__.Application           #type: Application
uidoc = __revit__.ActiveUIDocument      #type: UIDocument
doc   = revit.doc                       #type: Document
view  = revit.active_view
output = script.get_output()

# Main Code
# ==================================================
ducts = RevitDuct.from_selection(uidoc, doc, view)

if not ducts:
    forms.alert("Please select one or more ducts first.")
else:
    # keep both the ElementId and the weight
    weights = [(d.element.Id, d.id, d.length) 
            for d in ducts if d.length is not None]

    # Section title
    output.print_md("### Total Lengths")

    # Individual links with weights
    for eid, id_int, w in weights:
        output.print_md("- {}: {:.2f} ft".format(output.linkify(eid), w))

    # Select All link
    all_ids = List[ElementId]()
    for eid, _, _ in weights:
        all_ids.Add(eid)
    output.print_md("**{}**".format(output.linkify(all_ids)))

    # Footer total
    total = sum(w for _, _, w in weights)
    output.print_md("**➡️ Total Duct Length: {:.2f} ft**".format(total))