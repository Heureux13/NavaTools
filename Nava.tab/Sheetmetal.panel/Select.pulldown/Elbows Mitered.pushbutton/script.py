# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
import clr
from System.Collections.Generic import List
from revit_element import RevitElement
from revit_duct import RevitDuct, JointSize
from Autodesk.Revit.ApplicationServices import Application
from Autodesk.Revit.UI import UIDocument
from pyrevit import revit, forms, DB
from Autodesk.Revit.DB import *

# Button info
# # ======================================================================
__title__ = "Elbows Mitered"
__doc__ = """
****************************************************
Description:

Selects all 90° mitered elbows
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
sel_ducts = [d for d in ducts if d.family == "Square Bend"]

RevitElement.select_many(uidoc, sel_ducts)
forms.alert("Selected {} Mitered Elbows".format(len(sel_ducts)))
