# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
import sys
from Autodesk.Revit.DB import Transaction
from System.Collections.Generic import List
from revit_parameter import RevitParameter
from revit_element import RevitElement
from tag_duct import TagDuct
from revit_duct import RevitDuct, JointSize, CONNECTOR_THRESHOLDS
from revit_xyz import RevitXYZ
from Autodesk.Revit.ApplicationServices import Application
from Autodesk.Revit.UI import UIDocument
from pyrevit import revit, script, forms, DB
from Autodesk.Revit.DB import *
import clr

# Button display information
# =================================================
__title__ = "DO NOT PRESS"
__doc__ = """******************************************************************
Description:

Current goal fucntion of button is: select only spiral duct.

++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
This button is for testin code snippets and ideas before implementing them.
Odds are it will be constantly changing and not useful, its entire purpose
is for the author to have a quick button to test whatever code they are working on.
If you press it could do nothing, throw an error, or change something in your model.
Once working, it will most likely be moved to a more permanent location.
******************************************************************"""

# Variables
# ==================================================
app = __revit__.Application             # type: Application
uidoc = __revit__.ActiveUIDocument        # type: UIDocument
doc = revit.doc                         # type: Document
view = revit.active_view
output = script.get_output()

# Code
# ==================================================
ducts = RevitDuct.all(doc, view)

for d in ducts:
    if d.family and d.family.strip().lower() == "transition":
        print(d)

