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
from Autodesk.Revit.DB import Transaction, ElementId
from pyrevit import revit, forms, DB, script
from revit_duct import RevitDuct, JointSize, DuctAngleAllowance
from revit_xyz import RevitXYZ
from revit_tagging import RevitTagging
from revit_output import print_parameter_help
from revit_element import RevitElement

# Button info
# ==================================================
__title__ = "Short Joints"
__doc__ = """
Tag all short straight joints with length.
"""

# Variables
# ==================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
output = script.get_output()
view = revit.active_view
tagger = RevitTagging(doc=doc, view=view)

# View determination
# ==================================================
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

# Choose tag
# ==================================================
tag = tagger.get_label("_umi_length")

# Transaction
# ==================================================
already_tagged = []
needs_tagging = []
t = Transaction(doc, "Short Joints Tag")
t.Start()
try:
    # Begins our tagging process
    for d in fil_ducts:
        is_tagged = tagger.already_tagged(d.element, tag.Family.Name)
        if is_tagged:
            already_tagged.append(d)
            continue

        needs_tagging.append(d)
        ref, centroid = tagger.get_face_facing_view(d.element)
        if ref is not None and centroid is not None:
            tagger.place_tag(ref, tag, centroid)
        else:
            loc = d.element.Location
            if hasattr(loc, "Point") and loc.Point is not None:
                tagger.place_tag(d.element, tag, loc.Point)
            elif hasattr(loc, "Curve") and loc.Curve is not None:
                curve = loc.Curve
                midpoint = curve.Evaluate(0.5, True)
                tagger.place_tag(d.element, tag, midpoint)
            else:
                continue

    # Print already tagged list
    if already_tagged:
        # output.print_md("# Already Tagged: {}".format(len(already_tagged)))
        for i, d in enumerate(already_tagged, start=1):
            angle = RevitXYZ(d.element).straight_joint_degree()
            abs_angle = abs(angle)
            output.print_md(
                "### Index {} | Type: {} | Size: {} | Length: {} | Angle: {:.2f}° | Element ID: {}".format(
                    i, d.connector_0_type, d.size, d.length, abs_angle, output.linkify(
                        d.element.Id)
                )
            )

        element_ids = [d.element.Id for d in already_tagged]
        output.print_md(
            "# Total elements already tagged {}, {}".format(
                len(already_tagged), output.linkify(element_ids)
            )
        )

    output.print_md("---")

    if needs_tagging:
        # output.print_md("# Newly Tagged: {}".format(len(needs_tagging)))
        for i, d in enumerate(needs_tagging, start=1):
            angle = RevitXYZ(d.element).straight_joint_degree()
            abs_angle = abs(angle)
            output.print_md(
                "### Index {} | Type: {} | Size: {} | Length: {} | Angle: {:.2f}° | Element ID: {}".format(
                    i, d.connector_0_type, d.size, d.length, abs_angle, output.linkify(
                        d.element.Id)
                )
            )

        element_ids = [d.element.Id for d in needs_tagging]
        output.print_md(
            "# Total elements tagged {}, {}".format(
                len(needs_tagging), output.linkify(element_ids)
            )
        )

    output.print_md("---")

    element_ids = [d.element.Id for d in fil_ducts]
    output.print_md(
        "# Total elements {}, {}".format(
            len(fil_ducts), output.linkify(element_ids)
        )
    )

    # Final helper print
    print_parameter_help(output)

    t.Commit()
except Exception as e:
    output.print_md("Tag placement error: {}".format(e))
    t.RollBack()
    raise
