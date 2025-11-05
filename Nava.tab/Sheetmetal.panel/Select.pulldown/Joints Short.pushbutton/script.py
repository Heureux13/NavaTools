# -*- coding: utf-8 -*-
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""


__title__   = "Joints Short"
__doc__     = """
****************************************************************
Description:

Selects joints that are LESS than:
TDF     = 56.25"
S&D     = 59"
Spiral  = 120"
****************************************************************
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
fil_ducts  = [d for d in ducts if d.joint_size == JointSize.SHORT]

RevitElement.select_many(uidoc, fil_ducts)
forms.alert("Selected {} short joints".format(len(fil_ducts)))