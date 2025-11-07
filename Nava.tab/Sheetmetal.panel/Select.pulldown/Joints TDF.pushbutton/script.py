# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

__title__   = "Joints TDF"
__doc__     = """
************************************************************************
Description:

Selects all TDF joints
************************************************************************
"""

# Imports
# ==================================================
from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script, DB
from Autodesk.Revit.UI import UIDocument
from Autodesk.Revit.ApplicationServices import Application
from revit_duct import RevitDuct, JointSize, CONNECTOR_THRESHOLDS
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
output = script.get_output()

# Main Code
# ==================================================
allowed_joints = {
                    ("Straight", "TDC"),
                    ("Straight", "TDF"),
                    }

ducts = RevitDuct.all(doc, view)

valid_keys = set(CONNECTOR_THRESHOLDS.keys())

fil_ducts = [
    d for d in ducts if (d.family, d.connector_0) in allowed_joints
]

ids = List[ElementId]()
for d in fil_ducts:
    ids.Add(d.element.Id)

uidoc.Selection.SetElementIds(ids)

forms.alert("Selected {} TDF ducts".format(len(fil_ducts)))