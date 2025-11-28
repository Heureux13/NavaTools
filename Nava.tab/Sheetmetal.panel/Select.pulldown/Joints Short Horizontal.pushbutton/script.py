# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from Autodesk.Revit.DB import *
from pyrevit import revit, script
from revit_duct import RevitDuct, JointSize, DuctAngleAllowance, is_plan_view, is_section_view
from revit_xyz import RevitXYZ
from revit_element import RevitElement


__title__ = "Joints Short Horizontal"
__doc__ = """
Returns weight for duct(s) selected.
"""

# Variables
# ==================================================
app = __revit__.Application        # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc                    # type: Document
output = script.get_output()
view = revit.active_view

# Get all ducts
ducts = RevitDuct.all(doc, view)

# View/angle filter predicate
if is_plan_view(view):
    _vt = "plan"
elif is_section_view(view):
    _vt = "section"
else:
    _vt = "other"

# Helper function


def is_allowed(d):
    if d.joint_size != JointSize.SHORT:
        return False
    ang = RevitXYZ(d.element).straight_joint_degree()
    if isinstance(ang, (int, float)):
        a = abs(ang)
        if _vt == "plan" and DuctAngleAllowance.VERTICAL.contains(a):
            return False
        if _vt == "section" and DuctAngleAllowance.HORIZONTAL.contains(a):
            return False
        if _vt == "other":
            return False
    return True


# Filter down to short joints
fil_ducts = [d for d in ducts if is_allowed(d)]

# Start of select / print loop
if fil_ducts:
    RevitElement.select_many(uidoc, fil_ducts)
    output.print_md("# Selected {} short joints".format(len(fil_ducts)))
    output.print_md(
        "---")

    for i, fil in enumerate(fil_ducts, start=1):
        output.print_md(
            '### Index: {} | Type: {} | Size: {} | Length: {}" | ID: {}'.format(
                i, fil.connector_0_type, fil.size, fil.length, output.linkify(
                    fil.element.Id
                )
            )
        )

    element_ids = [d.element.Id for d in fil_ducts]
    output.print_md(
        "# Total elements {}, {}".format(
            len(
                element_ids
            ),
            output.linkify(
                element_ids
            )
        )
    )

    # Final print statements
    output.print_md(
        "---")
    output.print_md(
        "If info is missing, make sure you have the parameters turned on from Naviate")
    output.print_md(
        "All from Connectors and Fabrication, and size from Fab Properties")
else:
    output.print_md(
        "# There is nothing to select"
    )
