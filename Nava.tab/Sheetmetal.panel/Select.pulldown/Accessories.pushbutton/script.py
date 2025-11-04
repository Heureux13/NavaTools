# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

__title__   = "John's Button"
__doc__     = """
************************************************************************
Description:

Current goal fucntion of button is: Select boots per Goolsby request.

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

ducts = RevitDuct.all(doc, view)

conical         = [d for d in ducts if d.family == "ConicalTap - wDamper"]
boot_tap        = [d for d in ducts if d.family == "boot Tap - wDamper"]
long_coupler    = [d for d in ducts if d.family == "8inch Long Coupler wDamper"]

RevitElement.select_many(uidoc, conical + boot_tap + long_coupler)

forms.alert(
    "Selected {} conical taps\nSelected {} boot tap\nSelected {} long coupler".format(
        len(conical), len(boot_tap), len(long_coupler)
    )
)

print("It's called a do not press button for a reason, why did you press it?")