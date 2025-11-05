# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

__title__   = "DO NOT PRESS"
__doc__     = """
************************************************************************
Description:

Current goal fucntion of button is: return total weight including insulation.

++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
This button is for testin code snippets and ideas before implementing them.
Odds are it will be constantly changing and not useful, its entire purpose
is for the author to have a quick button to test whatever code they are working on.
If you press it could do nothing, throw an error, or change something in your model.
Once working, it will most likely be moved to a more permanent location.
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
    length = [(d.id, d.length) for d in ducts if d.length is not None]

    if not length:
        forms.alert("No weight data found for the selected ducts.")
    else:
        lines = ["Duct {}: {:.2f} ft".format(duct_id, l) for duct_id, l in length]
        total = sum(l for _, l in length)
        lines.append("----")
        lines.append("Total length: {:.2f} ft".format(total))

        print("\n".join(lines))