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
from Autodesk.Revit.DB import Transaction, ElementTransformUtils, XYZ, Line
from revit_duct import RevitDuct
from revit_xyz import RevitXYZ
from pyrevit import revit, script
import math

# Button display information
# =================================================
__title__ = "Tag Size"
__doc__ = """
Adds size tags to ducts
"""


# Code
# ==================================================
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

doc = revit.doc
view = revit.active_view
output = script.get_output()
ducts = RevitDuct.all(doc, view)
tagger = RevitTagging(doc, view)

needs_tagging = []
already_tagged = []
failed_to_tag = []  # Track elements that couldn't be tagged
skipped_unassigned = []  # Selected for review but not chosen this pass
skipped_vertical = []    # Skipped due to near-vertical orientation

straight_joint_families = {
    'straight',
    'spiral tube',
    'round duct',
    'tube',
    'spiral duct',
}

runs_to_skip = {
    "boot tap - wdamper",
    "boot saddle tap"
}

size_tags = {
    "_umi_size",
}

# Pre-fetch size tag family name for already-tagged checks
try:
    _size_tag_symbol = RevitTagging(doc, view).get_label("_umi_size")
    _size_tag_fam = (
        _size_tag_symbol.Family.Name if _size_tag_symbol and _size_tag_symbol.Family else "").strip()
except Exception:
    _size_tag_symbol = None
    _size_tag_fam = ""

t = Transaction(doc, "Tag Full Joints")
t.Start()
try:
    # Step 1: Gather all straight ducts in view
    all_ducts_in_view = ducts
    all_straights = [
        d for d in all_ducts_in_view
        if d.family and d.family.strip().lower() in straight_joint_families
    ]
    straight_lookup = {d.id: d for d in all_straights}
    elem_lookup = {d.id: d.element for d in all_straights}
    output.print_md(
        "**Found {} total straight ducts in view**".format(len(all_straights)))

    def build_tag_plan(straights, run_builder):
        processed_ids = set()
        selected_ids = set()
        considered_ids = set()
        run_infos = []
        run_count = 0

        while True:
            current = next(
                (s for s in straights if s.id not in processed_ids), None)
            if current is None:
                break

            run_count += 1
            try:
                run = run_builder(current, doc, view)
            except Exception as e:
                output.print_md(
                    "Warning: Could not build run from {}: {}".format(
                        current.element.Id.Value, str(e)))
                processed_ids.add(current.id)
                continue

            run_straights = [d for d in run if d.family and d.family.strip(
            ).lower() in straight_joint_families]
            run_fittings = [d for d in run if d.family and d.family.strip(
            ).lower() not in straight_joint_families]

            run_infos.append(
                (run_count, len(run_straights), len(run_fittings)))

            # Skip entire run if any fitting matches the skip list
            skip_run = any(
                (f.family or "").strip().lower() in runs_to_skip
                for f in run_fittings
            )
            if skip_run:
                for duct in run_straights:
                    considered_ids.add(duct.id)
                    processed_ids.add(duct.id)
                continue

            # Filter out straights under 24" long
            eligible_straights = [
                d for d in run_straights if d.length and d.length >= 24]

            # Separate already-tagged vs untagged for this run
            def has_size_tag(duct):
                try:
                    return tagger.already_tagged(duct.element, _size_tag_fam)
                except Exception:
                    return False

            already_tagged_run = [
                d for d in eligible_straights if has_size_tag(d)]
            untagged_straights = [
                d for d in eligible_straights if d not in already_tagged_run]

            total_eligible = len(eligible_straights)
            tag_these = []

            # Calculate how many total tags needed based on duct count
            # 1 tag for 1-8, 2 tags for 9-16, 3 tags for 17-24, etc.
            num_tags = max(1, int(math.ceil(float(total_eligible) / 8.0)))

            # If run already has enough BOD tags, skip adding more
            if len(already_tagged_run) >= num_tags:
                for duct in run_straights:
                    considered_ids.add(duct.id)
                    processed_ids.add(duct.id)
                continue

            remaining_needed = num_tags - len(already_tagged_run)
            n = len(untagged_straights)
            if n == 0:
                for duct in run_straights:
                    considered_ids.add(duct.id)
                    processed_ids.add(duct.id)
                continue

            chosen_idx = set()

            # First priority: find the longest untagged straight connected to any fitting
            longest_idx = -1
            longest_length = 0
            fitting_elem_ids = set(f.element.Id for f in run_fittings)

            for i, straight in enumerate(untagged_straights):
                try:
                    connected_elems = straight.get_connected_elements()
                    has_fitting = any(
                        e.Id in fitting_elem_ids
                        for e in connected_elems
                    )
                    if has_fitting and straight.length and straight.length > longest_length:
                        longest_idx = i
                        longest_length = straight.length
                except BaseException:
                    pass

            if longest_idx >= 0 and len(chosen_idx) < remaining_needed:
                chosen_idx.add(longest_idx)
                tag_these.append(untagged_straights[longest_idx])

            # Fill remaining slots with evenly distributed tags
            if len(chosen_idx) < remaining_needed and n > 1:
                still_needed = remaining_needed - len(chosen_idx)
                for i in range(still_needed):
                    idx = int(
                        round(i * (n - 1) / float(max(still_needed - 1, 1))))
                    if idx not in chosen_idx:
                        chosen_idx.add(idx)
                        tag_these.append(untagged_straights[idx])
            elif len(tag_these) == 0 and n >= 1:
                # Fallback: if no tags selected yet, tag the longest one
                longest_idx = max(
                    range(n), key=lambda i: untagged_straights[i].length) if n > 0 else 0
                tag_these.append(untagged_straights[longest_idx])

            for duct in tag_these:
                selected_ids.add(duct.id)
            for duct in run_straights:
                considered_ids.add(duct.id)
                processed_ids.add(duct.id)

        return selected_ids, considered_ids, run_infos, run_count

    size_selected_ids, size_considered_ids, size_run_infos, size_run_count = build_tag_plan(
        all_straights,
        RevitDuct.create_duct_run_same_height,
    )

    output.print_md("**Size runs: {} runs, {} straights chosen for size tags**".format(
        size_run_count, len(size_selected_ids)))

    # Build tag plan focused on size tags
    tag_plan = {}
    for eid in size_selected_ids:
        tag_plan[eid] = ['tag_size']

    connected_straights_to_tag = [elem_lookup[eid]
                                  for eid in tag_plan.keys() if eid in elem_lookup]

    considered_ids = size_considered_ids
    skipped_unassigned_ids = considered_ids - set(tag_plan.keys())
    skipped_unassigned = [straight_lookup[eid]
                          for eid in skipped_unassigned_ids if eid in straight_lookup]

    size_assign_count = sum(1 for v in tag_plan.values() if 'tag_size' in v)
    size_place_count = 0

    # Step 4: Tag the connected straights
    for elem in connected_straights_to_tag:
        elem_id = elem.Id.Value
        assignments = tag_plan.get(elem_id, [])

        if not assignments:
            skipped_unassigned.append(
                straight_lookup.get(elem_id, RevitDuct(doc, view, elem)))
            continue

        try:
            xyz_checker = RevitXYZ(elem)
            angle = xyz_checker.straight_joint_degree()
            if angle is not None and abs(angle) >= 85:
                skipped_vertical.append(
                    straight_lookup.get(elem_id, RevitDuct(doc, view, elem)))
                continue
        except BaseException:
            pass

        tagged_this_element = False
        had_existing = False
        tag_index = 0
        tag_spacing = 1.0

        for assignment in assignments:
            tags_to_use = size_tags
            placed = False
            had_existing_for_assignment = False

            for tag_name in tags_to_use:
                try:
                    tag_symbol = tagger.get_label(tag_name)
                    fam_name = (
                        tag_symbol.Family.Name if tag_symbol and tag_symbol.Family else "").strip().lower()
                    if not fam_name:
                        continue

                    if tagger.already_tagged(elem, fam_name):
                        had_existing_for_assignment = True
                        continue

                    face_ref, face_pt = tagger.get_face_facing_view(elem)
                    if face_ref is not None and face_pt is not None:
                        offset_pt = face_pt
                        try:
                            loc = elem.Location
                            if loc and hasattr(loc, 'Curve') and loc.Curve:
                                curve = loc.Curve
                                dir_vec = (curve.GetEndPoint(1) -
                                           curve.GetEndPoint(0)).Normalize()
                                offset_distance = tag_index * tag_spacing
                                offset_pt = XYZ(face_pt.X + dir_vec.X * offset_distance,
                                                face_pt.Y + dir_vec.Y * offset_distance,
                                                face_pt.Z + dir_vec.Z * offset_distance)
                        except BaseException:
                            pass

                        new_tag = tagger.place_tag(
                            face_ref, tag_symbol, offset_pt)
                        tagged_this_element = True
                        placed = True
                        tag_index += 1
                        size_place_count += 1

                        try:
                            loc = elem.Location
                            if loc and hasattr(loc, 'Curve') and loc.Curve:
                                curve = loc.Curve
                                dir_vec = (curve.GetEndPoint(1) -
                                           curve.GetEndPoint(0)).Normalize()
                                angle = math.atan2(dir_vec.Y, dir_vec.X)
                                axis = Line.CreateBound(offset_pt, XYZ(
                                    offset_pt.X, offset_pt.Y, offset_pt.Z + 1))
                                ElementTransformUtils.RotateElement(
                                    doc, new_tag.Id, axis, angle)
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
                                    dir_vec = (curve.GetEndPoint(1) -
                                               curve.GetEndPoint(0)).Normalize()
                                    offset_distance = tag_index * tag_spacing
                                    offset_pt = XYZ(center.X + dir_vec.X * offset_distance,
                                                    center.Y + dir_vec.Y * offset_distance,
                                                    center.Z + dir_vec.Z * offset_distance)
                            except BaseException:
                                pass

                            new_tag = tagger.place_tag(
                                elem, tag_symbol, offset_pt)
                            tagged_this_element = True
                            placed = True
                            tag_index += 1
                            size_place_count += 1

                            try:
                                loc = elem.Location
                                if loc and hasattr(loc, 'Curve') and loc.Curve:
                                    curve = loc.Curve
                                    dir_vec = (curve.GetEndPoint(1) -
                                               curve.GetEndPoint(0)).Normalize()
                                    angle = math.atan2(dir_vec.Y, dir_vec.X)
                                    axis = Line.CreateBound(offset_pt, XYZ(
                                        offset_pt.X, offset_pt.Y, offset_pt.Z + 1))
                                    ElementTransformUtils.RotateElement(
                                        doc, new_tag.Id, axis, angle)
                            except BaseException:
                                pass
                        else:
                            failed_to_tag.append(
                                (elem, "No valid placement location found"))

                    if placed:
                        break

                except Exception as e:
                    failed_to_tag.append((elem, str(e)))

            if not placed and had_existing_for_assignment:
                had_existing = True

        if tagged_this_element:
            needs_tagging.append(straight_lookup.get(
                elem_id, RevitDuct(doc, view, elem)))
        elif had_existing:
            already_tagged.append(straight_lookup.get(
                elem_id, RevitDuct(doc, view, elem)))
        else:
            failed_to_tag.append(
                (elem, "No tag created and no existing tag found"))
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
        output.print_md("## Already Tagged (has tag family on element)")
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

    if skipped_unassigned:
        output.print_md("## Skipped (not selected for tagging in this run)")
        for i, d in enumerate(skipped_unassigned, start=1):
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

    if skipped_vertical:
        output.print_md("## Skipped (near-vertical) ")
        for i, d in enumerate(skipped_vertical, start=1):
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
    # Diagnostic counters for tag set
    output.print_md(
        "# Size assigned/placed: {}/{}".format(size_assign_count, size_place_count))
    if already_tagged:
        already_ids = [d.element.Id for d in already_tagged]
        output.print_md("# Already tagged: {}, {}".format(
            len(already_tagged), output.linkify(already_ids)))
    if skipped_unassigned:
        skipped_ids = [d.element.Id for d in skipped_unassigned]
        output.print_md("# Skipped (not selected): {}, {}".format(
            len(skipped_unassigned), output.linkify(skipped_ids)))
    if skipped_vertical:
        vert_ids = [d.element.Id for d in skipped_vertical]
        output.print_md("# Skipped (vertical): {}, {}".format(
            len(skipped_vertical), output.linkify(vert_ids)))
    all_ducts = needs_tagging + already_tagged + skipped_unassigned + \
        [d for d, _ in failed_to_tag] + skipped_vertical
    all_ids = [d.element.Id for d in all_ducts]
    output.print_md("# Total: {}, {}".format(
        len(all_ducts), output.linkify(all_ids)))
    t.Commit()
except Exception as e:
    output.print_md("Tag placement error: {}".format(e))
    t.RollBack()
    raise
