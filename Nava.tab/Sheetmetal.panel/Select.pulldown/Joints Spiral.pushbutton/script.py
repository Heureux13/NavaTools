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
from revit_duct import RevitDuct, CONNECTOR_THRESHOLDS
from Autodesk.Revit.ApplicationServices import Application
from Autodesk.Revit.UI import UIDocument
from pyrevit import revit, forms, script
from Autodesk.Revit.DB import *
from System.Collections.Generic import List

# Button info
# ===================================================
__title__ = "Joint Spiral"
__doc__ = """
************************************************************************
Description:

Selects all spiral joints.
************************************************************************
"""

# Variables
# ==================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
view = revit.active_view
output = script.get_output()

# Main Code
# ==================================================
allowed_joints = {
    ("Tube", "GRC_Swage-Female"),
    ("Spiral Duct", "Raw"),
    ("Spiral Pipe", "Raw")
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

forms.alert("Selected {} spiral ducts".format(len(fil_ducts)))
