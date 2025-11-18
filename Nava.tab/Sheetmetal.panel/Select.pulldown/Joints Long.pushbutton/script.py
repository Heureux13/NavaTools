# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from revit_element import RevitElement
from revit_duct import DuctAngleAllowance, JointSize, RevitDuct
from Autodesk.Revit.ApplicationServices import Application
from Autodesk.Revit.UI import UIDocument
from pyrevit import revit, forms, DB
from Autodesk.Revit.DB import *

# Button info
# ===================================================
__title__ = "Joints Long"
__doc__ = """
****************************************************
Description:

Selects joints that are MORE than:
TDF     = 56"
S&D     = 59"
Spiral  = 120"
****************************************************
"""

# Variables
# ==================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
view = revit.active_view

# Main Code
# ==================================================

ducts = RevitDuct.all(doc, view)
fil_ducts = [d for d in ducts if d.joint_size == JointSize.LONG]

RevitElement.select_many(uidoc, fil_ducts)
forms.alert("Selected {} long joints".format(len(fil_ducts)))
