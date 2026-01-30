# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from revit_tagging import RevitTagging
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, IndependentTag, ElementTransformUtils, XYZ, Line
from revit_parameter import RevitParameter
from revit_duct import RevitDuct, JointSize
from revit_xyz import RevitXYZ
from pyrevit import revit, script
import math

# Button display information
# =================================================
__title__ = "Straight Duct Full"
__doc__ = """
Tags full straight duct connected to fittings with size label
"""


# Code
# ==================================================
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

doc = revit.doc
view = revit.active_view
output = script.get_output()
ducts = RevitDuct.all(doc, view)
rp = RevitParameter(doc, app)
tagger = RevitTagging(doc, view)

needs_tagging = []
already_tagged = []
failed_to_tag = []  # Track elements that couldn't be tagged

straight_joint_families = {
    'straight', 'spiral tube',
    'round duct', 'tube',
    'spiral duct',
}

tags = {
    "-FabDuct_ToD_MV_Tag",
    "-FabDuct_SIZE_MV_Tag",
}

t = Transaction(doc, "Tag Full Joints")
t.Start()
try:
    # Step 1: Get all ducts in view
    all_ducts_in_view = ducts
    output.print_md("**Found {} total ducts in view**".format(len(all_ducts_in_view)))

    # Step 2: Filter out straight families - get only non-straight (fittings, transitions, etc.)
    non_straight_ducts = [
        d for d in all_ducts_in_view if d.family and d.family.strip().lower() not in straight_joint_families
    ]
    output.print_md("**Found {} non-straight ducts (fittings/transitions)**".format(len(non_straight_ducts)))

    # Step 3: Find all straight ducts connected to those non-straight ducts
    connected_straights_to_tag = []
    seen_ids = set()

    for fitting in non_straight_ducts:
        for connector_idx in [0, 1, 2, 3]:
            connected_elems = fitting.get_connected_elements(connector_idx)
            for elem in connected_elems:
                if elem.Id.IntegerValue in seen_ids:
                    continue
                # Check if connected element is a straight duct
                connected_duct = RevitDuct(doc, view, elem)
                if connected_duct.family and connected_duct.family.strip().lower() in straight_joint_families:
                    connected_straights_to_tag.append(elem)
                    seen_ids.add(elem.Id.IntegerValue)

    output.print_md("**Found {} unique straight ducts connected to fittings**".format(len(connected_straights_to_tag)))

    # Step 4: Tag the connected straights, avoiding duplicates and skipping vertical ducts
    for elem in connected_straights_to_tag:
        # Skip if duct is vertical
        try:
            xyz_checker = RevitXYZ(elem)
            angle = xyz_checker.straight_joint_degree()
            if angle is not None and abs(angle) >= 85:  # Skip if near vertical (85° or more)
                already_tagged.append(RevitDuct(doc, view, elem))
                continue
        except BaseException:
            pass  # If we can't check angle, proceed with tagging

        existing_tag_fams = tagger.get_existing_tag_families(elem)
        tagged_this_element = False
        tag_index = 0
        tag_spacing = 1.0

        for tag_name in tags:
            try:
                tag_symbol = tagger.get_label(tag_name)
                fam_name = (tag_symbol.Family.Name if tag_symbol and tag_symbol.Family else "").strip().lower()
                if not fam_name:
                    continue

                # Skip if this tag family already exists on this element in this view
                if fam_name in existing_tag_fams:
                    continue

                face_ref, face_pt = tagger.get_face_facing_view(elem)
                if face_ref is not None and face_pt is not None:
                    # Calculate offset for multiple tags along duct direction
                    offset_pt = face_pt
                    try:
                        loc = elem.Location
                        if loc and hasattr(loc, 'Curve') and loc.Curve:
                            curve = loc.Curve
                            dir_vec = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize()
                            offset_distance = tag_index * tag_spacing
                            offset_pt = XYZ(face_pt.X + dir_vec.X * offset_distance,
                                            face_pt.Y + dir_vec.Y * offset_distance,
                                            face_pt.Z + dir_vec.Z * offset_distance)
                    except BaseException:
                        pass

                    new_tag = tagger.place_tag(face_ref, tag_symbol, offset_pt)
                    tagged_this_element = True
                    existing_tag_fams.add(fam_name)
                    tag_index += 1

                    # Rotate tag to match duct direction
                    try:
                        loc = elem.Location
                        if loc and hasattr(loc, 'Curve') and loc.Curve:
                            curve = loc.Curve
                            dir_vec = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize()
                            angle = math.atan2(dir_vec.Y, dir_vec.X)
                            axis = Line.CreateBound(offset_pt, XYZ(offset_pt.X, offset_pt.Y, offset_pt.Z + 1))
                            ElementTransformUtils.RotateElement(doc, new_tag.Id, axis, angle)
                    except BaseException:
                        pass
                else:
                    bbox = elem.get_BoundingBox(view)
                    if bbox is not None:
                        center = (bbox.Min + bbox.Max) / 2.0
                        offset_pt = center
                        try:
                            loc = elem.Location
                            if loc and hasattr(loc, 'Curve') and loc.Curve:
                                curve = loc.Curve
                                dir_vec = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize()
                                offset_distance = tag_index * tag_spacing
                                offset_pt = XYZ(center.X + dir_vec.X * offset_distance,
                                                center.Y + dir_vec.Y * offset_distance,
                                                center.Z + dir_vec.Z * offset_distance)
                        except BaseException:
                            pass

                        new_tag = tagger.place_tag(elem, tag_symbol, offset_pt)
                        tagged_this_element = True
                        existing_tag_fams.add(fam_name)
                        tag_index += 1

                        # Rotate tag to match duct direction
                        try:
                            loc = elem.Location
                            if loc and hasattr(loc, 'Curve') and loc.Curve:
                                curve = loc.Curve
                                dir_vec = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize()
                                angle = math.atan2(dir_vec.Y, dir_vec.X)
                                axis = Line.CreateBound(offset_pt, XYZ(offset_pt.X, offset_pt.Y, offset_pt.Z + 1))
                                ElementTransformUtils.RotateElement(doc, new_tag.Id, axis, angle)
                        except BaseException:
                            pass
                    else:
                        failed_to_tag.append((elem, "No valid placement location found"))
                        continue
            except Exception as e:
                failed_to_tag.append((elem, str(e)))
                continue

        # Add element to appropriate list
        if tagged_this_element:
            needs_tagging.append(RevitDuct(doc, view, elem))
        else:
            already_tagged.append(RevitDuct(doc, view, elem))
    output.print_md("---")

    if failed_to_tag:
        output.print_md("## Failed to Tag")
        for i, (elem, reason) in enumerate(failed_to_tag, start=1):
            output.print_md("### Index {} | Element ID: {} | Reason: {}".format(
                i, output.linkify(elem.Id), reason))
        output.print_md("---")

    if needs_tagging:
        output.print_md("## Newly Tagged")
        for i, d in enumerate(needs_tagging, start=1):
            output.print_md(
                "### No.{} | ID: {} | Fam: {} | Size: {} | Le: {:06.2f}".format(
                    i,
                    output.linkify(d.element.Id),
                    d.family,
                    d.size,
                    d.length
                )
            )
        output.print_md("---")

    if already_tagged:
        output.print_md("## Already Tagged")
        for i, d in enumerate(already_tagged, start=1):
            output.print_md(
                "### Index {} | Size: {} | Family: {} | Length: {:06.2f} | Element ID: {}".format(
                    i,
                    d.size,
                    d.family,
                    d.length,
                    output.linkify(d.element.Id)
                )
            )
        output.print_md("---")

    if needs_tagging:
        newly_ids = [d.element.Id for d in needs_tagging]
        output.print_md("# Newly tagged: {}, {}".format(
            len(needs_tagging), output.linkify(newly_ids)))
    if already_tagged:
        already_ids = [d.element.Id for d in already_tagged]
        output.print_md("# Already tagged: {}, {}".format(
            len(already_tagged), output.linkify(already_ids)))
    all_ducts = needs_tagging + already_tagged
    all_ids = [d.element.Id for d in all_ducts]
    output.print_md("# Total: {}, {}".format(
        len(all_ducts), output.linkify(all_ids)))
    t.Commit()
except Exception as e:
    output.print_md("Tag placement error: {}".format(e))
    t.RollBack()
    raise
