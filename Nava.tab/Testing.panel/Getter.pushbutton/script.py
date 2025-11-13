# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from System.Collections.Generic import List
from Autodesk.Revit.ApplicationServices import Application
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import UIDocument
from pyrevit import revit, forms, DB, script
from revit_element import RevitElement
from revit_duct import RevitDuct, JointSize, DuctAngleAllowance
from revit_xyz import RevitXYZ
import clr

# Button display information
# =================================================
__title__ = "Getter"
__doc__ = """
____________________________________________________
Description:

Returns the value of whatever is coded here, should be a parameter or
something simple
____________________________________________________
"""

# Variables
# ==================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Documents  
view = revit.active_view
output = script.get_output()

# Main Code
# ==================================================
# Get the current view object
view = revit.active_view  # or uidoc.ActiveView

# Determine view type for filtering
if view.ViewType == DB.ViewType.FloorPlan:
    current_view_type = "floor"
elif view.ViewType == DB.ViewType.Section:
    current_view_type = "section"
else:
    current_view_type = "other"

ducts = RevitDuct.all(doc, view)

if not ducts:
    forms.alert("Please select a duct element", exitscript=True)

# Example: filter short joints and by angle depending on view type
from revit_duct import DuctAngleAllowance

fil_ducts = []
 
for d in ducts:
    if d.joint_size == JointSize.SHORT:
        xyz = RevitXYZ(d.element)
        angle = xyz.straight_joint_degree()
        output.print_md("Duct ID: {}, Angle: {}".format(d.id, angle))
        if isinstance(angle, (int, float)):
            abs_angle = abs(angle)
            if current_view_type == "floor":
                # Omit vertical ducts in plan view
                if not DuctAngleAllowance.VERTICAL.contains(abs_angle):
                    fil_ducts.append(d)
            elif current_view_type == "section":
                # Omit horizontal ducts in section view
                if not DuctAngleAllowance.HORIZONTAL.contains(abs_angle):
                    fil_ducts.append(d)
            else:
                fil_ducts.append(d)
        else:
            fil_ducts.append(d)

RevitElement.select_many(uidoc, fil_ducts)
output.print_md("Selected {} short joints 1".format(len(fil_ducts)))