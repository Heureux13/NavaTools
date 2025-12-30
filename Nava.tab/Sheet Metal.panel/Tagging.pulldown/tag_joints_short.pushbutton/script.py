# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from Autodesk.Revit.DB import Transaction
from pyrevit import revit, forms, DB, script
from revit_duct import RevitDuct, JointSize, DuctAngleAllowance
from revit_xyz import RevitXYZ
from revit_tagging import RevitTagging
from revit_output import print_disclaimer

# Button info
# ==================================================
__title__ = "Straight Duct Short"
__doc__ = """
Tag all short straight duct with length.
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
skipped_ducts = []
for d in ducts:
    if d.joint_size != JointSize.SHORT:
        continue
    angle = RevitXYZ(d.element).straight_joint_degree()
    if not isinstance(angle, (int, float)):
        skipped_ducts.append((d, "angle_not_numeric", angle))
        continue

    abs_angle = abs(angle)
    skip_reason = None
    if current_view_type == "floor":
        if DuctAngleAllowance.VERTICAL.contains(abs_angle):
            skip_reason = "vertical_angle_in_floor_view"
    elif current_view_type == "section":
        if DuctAngleAllowance.HORIZONTAL.contains(abs_angle):
            skip_reason = "horizontal_angle_in_section_view"

    if skip_reason:
        skipped_ducts.append((d, skip_reason, abs_angle))
        continue

    fil_ducts.append(d)

# Choose tag
# ==================================================
tag = tagger.get_label("-FabDuct_LENGTH_FIX_Tag")


def _below_bbox_point(elem, base_point=None, offset_z=-0.5):
    """Return a point directly below the element bbox (world Z)."""
    try:
        bbox = elem.get_BoundingBox(view)
        if bbox is None:
            return None
        bx = base_point.X if base_point else (bbox.Min.X + bbox.Max.X) / 2.0
        by = base_point.Y if base_point else (bbox.Min.Y + bbox.Max.Y) / 2.0
        return DB.XYZ(bx, by, bbox.Min.Z + float(offset_z))
    except Exception:
        return None


def _has_tag_type(elem, tag_symbol):
    """Return True if elem already has a tag of the same type as tag_symbol in this view."""
    try:
        if elem is None or tag_symbol is None:
            return False
        target_type_id = getattr(tag_symbol, "Id", None)
        if not target_type_id:
            return False
        tags = list(
            DB.FilteredElementCollector(doc, view.Id)
            .OfClass(DB.IndependentTag)
            .ToElements()
        )
        for itag in tags:
            try:
                tagged_ids = None
                if hasattr(itag, "GetTaggedLocalElementIds"):
                    tagged_ids = itag.GetTaggedLocalElementIds() or []
                elif hasattr(itag, "TaggedLocalElementId"):
                    tagged_ids = [itag.TaggedLocalElementId]
                if not tagged_ids or elem.Id not in tagged_ids:
                    continue
                if itag.GetTypeId() == target_type_id:
                    return True
            except Exception:
                continue
    except Exception:
        return False
    return False


# Transaction
# ==================================================
already_tagged = []
needs_tagging = []
t = Transaction(doc, "Short Joints Tag")
t.Start()
try:
    # Begins our tagging process
    for d in fil_ducts:
        is_tagged = _has_tag_type(d.element, tag) or tagger.already_tagged(d.element, tag.Family.Name)
        if is_tagged:
            already_tagged.append(d)
            continue

        needs_tagging.append(d)
        ref, centroid = tagger.get_face_facing_view(d.element)
        if ref is not None and centroid is not None:
            pt = _below_bbox_point(d.element, centroid) or centroid
            tagger.place_tag(ref, tag, pt)
        else:
            loc = d.element.Location
            if hasattr(loc, "Point") and loc.Point is not None:
                pt = _below_bbox_point(d.element, loc.Point) or loc.Point
                tagger.place_tag(d.element, tag, pt)
            elif hasattr(loc, "Curve") and loc.Curve is not None:
                curve = loc.Curve
                midpoint = curve.Evaluate(0.25, True)
                pt = _below_bbox_point(d.element, midpoint) or midpoint
                tagger.place_tag(d.element, tag, pt)
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
                    i,
                    d.connector_0_type,
                    d.size,
                    d.length,
                    abs_angle,
                    output.linkify(d.element.Id)
                )
            )

        element_ids = [d.element.Id for d in already_tagged]
        output.print_md(
            "# Total elements already tagged {}, {}".format(
                len(already_tagged),
                output.linkify(element_ids)
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
                    i,
                    d.connector_0_type,
                    d.size,
                    d.length,
                    abs_angle,
                    output.linkify(d.element.Id)
                )
            )

        element_ids = [d.element.Id for d in needs_tagging]
        output.print_md(
            "# Total elements tagged {}, {}".format(
                len(needs_tagging),
                output.linkify(element_ids)
            )
        )

    output.print_md("---")

    element_ids = [d.element.Id for d in fil_ducts]
    output.print_md(
        "# Total elements {}, {}".format(
            len(fil_ducts),
            output.linkify(element_ids)
        )
    )

    if skipped_ducts:
        output.print_md("---")
        output.print_md("# Skipped Ducts: {}".format(len(skipped_ducts)))
        for d, reason, value in skipped_ducts:
            output.print_md(
                "### Reason: {} | Value: {} | Element ID: {}".format(
                    reason,
                    value,
                    output.linkify(d.element.Id)
                )
            )

    # Final helper print
    print_disclaimer(output)

    t.Commit()
except Exception as e:
    output.print_md("Tag placement error: {}".format(e))
    t.RollBack()
    raise
