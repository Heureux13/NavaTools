# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

import clr
from System.Collections.Generic import List
from revit_element import RevitElement
from tag_duct import TagDuct
from revit_duct import RevitDuct, JointSize
from Autodesk.Revit.ApplicationServices import Application
from Autodesk.Revit.UI import UIDocument
from pyrevit import revit, forms, DB
from Autodesk.Revit.DB import *
__title__ = "Getter"
__doc__ = """
____________________________________________________
Description:

Returns the value of whatever is coded here, should be a parameter or
something simple
____________________________________________________
"""

# Imports
# ==================================================

# .NET Imports
# ==================================================

# Variables
# ==================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
view = revit.active_view

# Main Code
# ==================================================
ducts = RevitDuct.from_selection(uidoc, doc)

if not ducts:
    forms.alert("Please select one or more ducts first")
else:
    values = [str(d.weight_total) for d in ducts]
    forms.alert("\n".join(values))
