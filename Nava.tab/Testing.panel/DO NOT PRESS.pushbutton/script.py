# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from pyrevit import revit, script
from revit_duct import RevitDuct, JointSize
from revit_parameter import RevitParameter
from revit_tagging import RevitTagging
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, IndependentTag, ElementId

# Button display information
# =================================================
__title__ = "DO NOT PRESS"
__doc__ = """
******************************************************************
Gives offset information about specific duct fittings
******************************************************************
"""

# Variables
# ==================================================
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
view = revit.active_view
output = script.get_output()
ducts = RevitDuct.all(doc, view)
# ducts = RevitDuct.from_selection(uidoc, doc)
rp = RevitParameter(doc, app)


# Code
# ==================================================
with Transaction(doc, "Offset Parameter") as t:
    t.Start()

    # Step 1: Filter to only non-straight ducts (fittings, transitions, etc.)
    non_straight_ducts = [
        d for d in ducts if d.family and d.family.strip().lower() not in [
            "straight", "spiral tube", "round duct", "tube"]]
    output.print_md("**Found {} non-straight ducts**".format(len(non_straight_ducts)))

    # Step 2: Get all straight ducts connected to non-straight ducts
    connected_straights = []
    seen_ids = set()

    for d in non_straight_ducts:
        # Check all connectors (0, 1, 2)
        for connector_index in [0, 1, 2]:
            connected_elements = d.get_connected_elements(connector_index)
            for elem in connected_elements:
                # Skip if already processed
                if elem.Id.IntegerValue in seen_ids:
                    continue

                # Wrap as RevitDuct to check properties
                connected_duct = RevitDuct(doc, view, elem)

                # Only add if it's a straight or spiral duct
                if connected_duct.family and connected_duct.family.strip().lower(
                ) in ["straight", "spiral tube", "round duct", "tube"]:
                    connected_straights.append(elem)
                    seen_ids.add(elem.Id.IntegerValue)

    output.print_md("**Found {} unique connected straight ducts**".format(len(connected_straights)))

    # Step 3: Filter to only full-size joints
    full_joint_straights = []
    for elem in connected_straights:
        duct = RevitDuct(doc, view, elem)
        if duct.joint_size == JointSize.FULL:
            full_joint_straights.append(elem)

    output.print_md("**Found {} full-size joints to tag**".format(len(full_joint_straights)))

    # Step 4: Tag the full-size joints
    tagger = RevitTagging(doc, view)
    tagged_count = 0
    skipped_count = 0

    # Pre-build a set of (element_id, tag_family_name) tuples for already tagged elements
    already_tagged_set = set()
    existing_tags = FilteredElementCollector(doc, view.Id).OfClass(IndependentTag).ToElements()

    for tag in existing_tags:
        try:
            # GetTaggedReferences returns a Python list, not a .NET collection
            refs = tag.GetTaggedReferences()
            if refs and len(refs) > 0:
                # Get the first reference and extract its ElementId
                ref = refs[0]
                tagged_elem_id = ref.ElementId

                # Get the tag family
                tag_type = doc.GetElement(tag.GetTypeId())
                if tag_type and hasattr(tag_type, 'Family'):
                    already_tagged_set.add((tagged_elem_id.IntegerValue, tag_type.Family.Name))
        except BaseException:
            pass

    output.print_md("**Found {} existing tags in view**".format(len(already_tagged_set)))

    for elem in full_joint_straights:
        tags = ["0_bod", "0_size"]
        for tag_name in tags:
            try:
                tag_symbol = tagger.get_label(tag_name)

                # Check if already tagged using pre-built set
                if (elem.Id.IntegerValue, tag_symbol.Family.Name) in already_tagged_set:
                    skipped_count += 1
                    continue

                # Get face reference and centroid for fabrication duct
                face_ref, face_pt = tagger.get_face_facing_view(elem)

                if face_ref is not None and face_pt is not None:
                    tagger.place_tag(face_ref, tag_symbol, face_pt)
                    tagged_count += 1
                    # Add to set so we don't double-tag in this run
                    already_tagged_set.add((elem.Id.IntegerValue, tag_symbol.Family.Name))
                else:
                    # Fallback to bounding box center
                    bbox = elem.get_BoundingBox(view)
                    if bbox is not None:
                        center = (bbox.Min + bbox.Max) / 2.0
                        tagger.place_tag(elem, tag_symbol, center)
                        tagged_count += 1
                        # Add to set so we don't double-tag in this run
                        already_tagged_set.add((elem.Id.IntegerValue, tag_symbol.Family.Name))

            except Exception as e:
                output.print_md("  Failed to tag element {}: {}".format(
                    elem.Id.IntegerValue, str(e)))

    output.print_md("\n**Tagged {} elements, skipped {} already tagged**".format(tagged_count, skipped_count))

    t.Commit()

    output.print_md("✓ Complete")
