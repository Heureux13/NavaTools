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
from pyrevit import revit, forms, DB
from Autodesk.Revit.DB import *

# Button info
# ===================================================
__title__ = "Joints S&D"
__doc__ = """
******************************************************************
Description:

Selects all S&D joints
******************************************************************
"""

# Variables
# ==================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
view = revit.active_view

# Main Code
# ==================================================
allowed_joints = {
    ("Straight", "Slip & Drive"),
    ("Straight", "S&D"),
    ("Straight", "Standing S&D")
}

ducts = RevitDuct.all(doc, view)

valid_keys = set(CONNECTOR_THRESHOLDS.keys())

fil_ducts = [
    d for d in ducts if (d.family, d.connector_0) in allowed_joints
]

RevitElement.select_many(uidoc, fil_ducts)
forms.alert("Selected {} S&D ducts.".format(len(fil_ducts)))
