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
from revit_duct import RevitDuct, JointSize, DuctAngleAllowance
from revit_xyz import RevitXYZ


__title__ = "Joints Short Horizontal"

# Variables
# ==================================================
app = __revit__.Application        # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc                    # type: Document
output = script.get_output()

# View determination
# ==================================================
view = revit.active_view
if view.ViewType == DB.ViewType.FloorPlan:
    current_view_type = "floor"
elif view.ViewType == DB.ViewType.Section:
    current_view_type = "section"
else:
    current_view_type = "other"

# Collect ducts in view
# ==================================================
ducts = RevitDuct.all(doc, view)
if not ducts:
    forms.alert("No ducts found in the current view", exitscript=True)

# Filtered results
# ==================================================
fil_ducts = []
for d in ducts:
    if d.joint_size != JointSize.SHORT:
        continue
    angle = RevitXYZ(d.element).straight_joint_degree()
    if isinstance(angle, (int, float)):
        abs_angle = abs(angle)
        if current_view_type == "floor":
            if DuctAngleAllowance.VERTICAL.contains(abs_angle):
                continue
        elif current_view_type == "section":
            if DuctAngleAllowance.HORIZONTAL.contains(abs_angle):
                continue
    fil_ducts.append(d)

# Select the filtered ducts
# ==================================================
if fil_ducts:
    all_ids = List[ElementId]()
    for d in fil_ducts:
        all_ids.Add(d.element.Id)
    uidoc.Selection.SetElementIds(all_ids)
else:
    uidoc.Selection.SetElementIds(List[ElementId]())

# Out put results
# ==================================================
output.print_md("## Selected {} short joint(s)".format(len(fil_ducts)))
output.print_md("---")

for d in fil_ducts:
    eid = d.element.Id
    angle = RevitXYZ(d.element).straight_joint_degree()
    abs_angle = abs(angle)

    link_token = output.linkify(eid)

    if link_token is None:
        if isinstance(abs_angle, (int, float)):
            output.print_md("  Angle: {:.2f}°".format(abs_angle))
        else:
            output.print_md("  Angle: N/A")
        output.print_md("------------------------------------------------")
    else:
        # link_token is printable inline
        if isinstance(abs_angle, (int, float)):
            output.print_md(
                "- {}  |  Angle: {:.2f}°".format(link_token, abs_angle))
        else:
            output.print_md("- {}  |  Angle: N/A".format(link_token))

# Add a single "Select all" link (optional)
if fil_ducts:
    output.print_md("**{}**".format(output.linkify(all_ids)))
