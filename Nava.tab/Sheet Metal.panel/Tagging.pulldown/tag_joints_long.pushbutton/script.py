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
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, IndependentTag
from revit_parameter import RevitParameter
from revit_duct import RevitDuct, JointSize
from pyrevit import revit, script

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

tagged_count = 0
skipped_count = 0

straight_joint_families = {
    'straight', 'spiral tube',
    'round duct', 'tube',
    'spiral duct',
}

tags = {
    "_umi_bod",
    "_umi_size"
}

t = Transaction(doc, "Tag Full Joints")
t.Start()
try:
    # Step 1: Filter to only non-straight ducts (fittings, transitions, etc.)
    non_straight_ducts = [
        d for d in ducts if d.family and d.family.strip().lower() not in straight_joint_families
    ]
    output.print_md(
        "**Found {} non-straight ducts**".format(len(non_straight_ducts)))

    # Step 2: Get all straight ducts connected to non-straight ducts
    connected_straights = []
    seen_ids = set()
    for d in non_straight_ducts:
        for connector_index in [0, 1, 2]:
            connected_elements = d.get_connected_elements(connector_index)
            for elem in connected_elements:
                if elem.Id.IntegerValue in seen_ids:
                    continue
                connected_duct = RevitDuct(doc, view, elem)
                if connected_duct.family and connected_duct.family.strip(
                ).lower() in straight_joint_families:
                    connected_straights.append(elem)
                    seen_ids.add(elem.Id.IntegerValue)
    output.print_md(
        "**Found {} unique connected straight ducts**".format(len(connected_straights)))

    # Step 3: Filter to only full-size joints
    full_joint_straights = []
    for elem in connected_straights:
        duct = RevitDuct(doc, view, elem)
        if duct.joint_size == JointSize.FULL:
            full_joint_straights.append(elem)
    output.print_md(
        "**Found {} full-size joints to tag**".format(len(full_joint_straights)))

    # Step 4: Tag the full-size joints
    already_tagged_set = set()
    from Autodesk.Revit.DB import FilteredElementCollector, IndependentTag
    existing_tags = FilteredElementCollector(
        doc, view.Id).OfClass(IndependentTag).ToElements()
    for tag in existing_tags:
        try:
            refs = tag.GetTaggedReferences()
            if refs and len(refs) > 0:
                ref = refs[0]
                tagged_elem_id = ref.ElementId
                tag_type = doc.GetElement(tag.GetTypeId())
                if tag_type and hasattr(tag_type, 'Family'):
                    already_tagged_set.add(
                        (tagged_elem_id.IntegerValue, tag_type.Family.Name))
        except BaseException:
            pass
    output.print_md(
        "**Found {} existing tags in view**".format(len(already_tagged_set)))

    for elem in full_joint_straights:
        for tag_name in tags:
            try:
                tag_symbol = tagger.get_label(tag_name)
                if (elem.Id.IntegerValue, tag_symbol.Family.Name) in already_tagged_set:
                    skipped_count += 1
                    continue
                face_ref, face_pt = tagger.get_face_facing_view(elem)
                if face_ref is not None and face_pt is not None:
                    tagger.place_tag(face_ref, tag_symbol, face_pt)
                    tagged_count += 1
                    already_tagged_set.add(
                        (elem.Id.IntegerValue, tag_symbol.Family.Name))
                else:
                    bbox = elem.get_BoundingBox(view)
                    if bbox is not None:
                        center = (bbox.Min + bbox.Max) / 2.0
                        tagger.place_tag(elem, tag_symbol, center)
                        tagged_count += 1
                        already_tagged_set.add(
                            (elem.Id.IntegerValue, tag_symbol.Family.Name))
                    else:
                        skipped_count += 1
                        continue
            except Exception as e:
                output.print_md("Tag placement error: {}".format(e))
                skipped_count += 1
                continue
    output.print_md("---")
    output.print_md("# Tagging Summary")
    output.print_md("**Tagged: {}**".format(tagged_count))
    output.print_md(
        "**Skipped (already tagged or no placement): {}**".format(skipped_count))
    output.print_md(
        "**Total full joints processed: {}**".format(len(full_joint_straights)))
    t.Commit()
except Exception as e:
    output.print_md("Tag placement error: {}".format(e))
    t.RollBack()
    raise
