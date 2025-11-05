# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

__title__   = "Weight Total"
__doc__     = """
************************************************************************
Description:

Returns metal duct weight + insulation duct weight
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
ducts = RevitDuct.from_selection(uidoc, doc, view)

if not ducts:
    forms.alert("Please select one or more ducts first.")
else:
    weights = [(d.id, d.weigth_total) for d in ducts if d.weigth_total is not None]

    if not weights:
        forms.alert("No weight data found for the selected ducts.")
    else:
        lines = ["Duct {}: {:.2f} lbs".format(duct_id, w) for duct_id, w in weights]
        total = sum(w for _, w in weights)
        lines.append("----")
        lines.append("Total weight: {:.2f} lbs".format(total))

        forms.alert("\n".join(lines))