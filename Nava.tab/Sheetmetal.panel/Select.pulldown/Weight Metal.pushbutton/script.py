# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

__title__   = "Weight Metal"
__doc__     = """
************************************************************************
Description:

Gets the weight of whatever duct is selected, add them if more than one
************************************************************************
"""

# Imports
# ==================================================
from Autodesk.Revit.DB import *
from pyrevit import revit, forms, DB
from Autodesk.Revit.UI import UIDocument
from Autodesk.Revit.ApplicationServices import Application
from revit_duct import RevitDuct, JointSize
from tag_duct import TagDuct
from revit_element import RevitElement

#.NET Imports
# ==================================================
import clr
clr.AddReference('System')
from System.Collections.Generic import List


# Variables
# ==================================================
app   = __revit__.Application           #type: Application
uidoc = __revit__.ActiveUIDocument      #type: UIDocument
doc   = revit.doc                       #type: Document
view  = revit.active_view

# Main Code
# ==================================================
sel_ids = RevitDuct.from_selection(uidoc, doc)

if not sel_ids:
    forms.alert("Please select one or more ducts first.")
else:
    ducts = []
    for elid in sel_ids:
        el = doc.GetElement(elid)
        ducts.append(RevitDuct(doc, view, el))

    # Collect weights
    weights = [(d.id, d.weight_metal) for d in ducts if d.weight_metal]

    if not weights:
        forms.alert("No weight data found for the selected ducts.")
    else:
        # Build a message with each duct’s weight
        lines = ["Duct {}: {} lbs".format(duct_id, w) for duct_id, w in weights]
        total = sum(w for _, w in weights)
        lines.append("----")
        lines.append("Total weight: {} lbs".format(total))

        forms.alert("\n".join(lines))