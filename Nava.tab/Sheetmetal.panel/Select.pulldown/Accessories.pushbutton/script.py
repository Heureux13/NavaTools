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
from System.Collections.Generic import List
import clr


# Variables
# ==================================================
app   = __revit__.Application           #type: Application
uidoc = __revit__.ActiveUIDocument      #type: UIDocument
doc   = revit.doc                       #type: Document
view  = revit.active_view

# Main Code
# ==================================================

ducts = RevitDuct.all(doc, view)

conical         = [d for d in ducts if d.family.lower().strip() == "conicaltap - wdamper"]
boot_tap        = [d for d in ducts if d.family.lower().strip() == "boot tap - wdamper"]
long_coupler    = [d for d in ducts if d.family.lower().strip() == "8inch long coupler wdamper"]
end_cap         = [d for d in ducts if d.family.lower().strip() == "cap"]

RevitElement.select_many(uidoc, conical + boot_tap + long_coupler + end_cap)

forms.alert(
    "Selected {} conical taps\n"
    "Selected {} boot tap\n"
    "Selected {} long coupler\n"
    "Selected {} end caps".format(
        len(conical), len(boot_tap), len(long_coupler), len(end_cap)
    )
)
