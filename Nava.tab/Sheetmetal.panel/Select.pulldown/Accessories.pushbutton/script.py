# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

__title__   = "Accessories"
__doc__     = """
************************************************************************
Description:

Selects all end caps and taps
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
end_cap    = [d for d in ducts if d.family == "cap"]

RevitElement.select_many(uidoc, conical + boot_tap + long_coupler + end_cap)

forms.alert(
    "Selected {} conical taps\nSelected {} boot tap\nSelected {} long coupler\nSelected {} end caps".format(
        len(conical), len(boot_tap), len(long_coupler), len(end_cap)
    )
)
