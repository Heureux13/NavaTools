# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

__title__   = "Elbows Radius"
__doc__     = """
************************************************************************
Description:

Select all mitered elbows not 90° and all radius elbows.

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
rad_ducts   = [d for d in ducts if d.family == "Radius Bend"]
el_ducts    = [d for d in ducts if d.family == "Elbow"]

RevitElement.select_many(uidoc, el_ducts + rad_ducts)
forms.alert("Selected {} radius elbows\nSelected {} Mitered elbows not 90°".format(len(rad_ducts), len(el_ducts)))