# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from tagging.revit_tagging import RevitTagging
from Autodesk.Revit.DB import Transaction, ElementTransformUtils, XYZ, Line
from ducts.revit_duct import (
    RevitDuct,
    JointSize,
    CONNECTOR_THRESHOLDS,
    DEFAULT_SHORT_THRESHOLD_IN,
)
from ducts.revit_xyz import RevitXYZ
from tagging.tag_config import (
    DEFAULT_TAG_SLOT_CANDIDATES,
    STRAIGHT_JOINT_FAMILIES,
    DEFAULT_JOINT_TAG_SLOTS,
    SLOT_BOD,
    SLOT_LENGTH,
    SLOT_SIZE,
)
from pyrevit import revit, script
import math

# Button display information
# =================================================
__title__ = "Tag Joints Full"
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
tagger = RevitTagging(doc, view)

needs_tagging = []
already_tagged = []
skipped_vertical = []
no_tag_needed = []
failed_to_tag = []  # Track elements that couldn't be tagged
non_fatal_warnings = []
missing_tag_labels = set()
slot_resolution_cache = {}

TAG_SLOT_CANDIDATES = {
    slot: list(candidates)
    for slot, candidates in DEFAULT_TAG_SLOT_CANDIDATES.items()
}

straight_joint_families = set(STRAIGHT_JOINT_FAMILIES)

full_run_tag_slots = list(DEFAULT_JOINT_TAG_SLOTS)
SHORT_LENGTH_ONLY_RATIO = 0.25
SPIRAL_NO_TAG_MAX_LENGTH_IN = 12.0
PRINT_RUN_ASSIGNMENT_DEBUG = True


def get_element_id_value(elem):
    try:
        return elem.Id.Value
    except Exception:
        return elem.Id.IntegerValue


def resolve_tag_slot(tag_or_slot_name):
    """Resolve a slot key (or literal tag name) to a loaded tag symbol.

    Returns tuple: (tag_symbol, family_name_lower, matched_candidate_name)
    or (None, None, None) when no candidate exists in current project.
    """
    key = str(tag_or_slot_name or '').strip()
    if not key:
        return None, None, None

    cache_key = key.upper()
    if cache_key in slot_resolution_cache:
        return slot_resolution_cache[cache_key]

    # If key is not a configured slot, treat it as a literal candidate.
    candidates = TAG_SLOT_CANDIDATES.get(cache_key)
    if not candidates:
        candidates = [key]

    resolved = (None, None, None)
    seen = set()
    attempted = []
    for candidate in candidates:
        name = str(candidate or '').strip()
        if not name:
            continue
        if name.lower() in seen:
            continue
        seen.add(name.lower())
        attempted.append(name)
        try:
            tag_symbol = tagger.get_label(name)
        except LookupError:
            tag_symbol = None

        if tag_symbol is None:
            continue

        fam_name = (
            tag_symbol.Family.Name if tag_symbol and tag_symbol.Family else ""
        ).strip().lower()
        resolved = (tag_symbol, fam_name, name)
        break

    if resolved[0] is None:
        missing_tag_labels.update(attempted)

    slot_resolution_cache[cache_key] = resolved
    return resolved


def is_straight_family(family_name):
    return bool(
        family_name and family_name.strip().lower() in straight_joint_families
    )


def get_straight_neighbor_ids(duct_wrap, straight_ids):
    def connector_sort_key(connector):
        try:
            origin = connector.Origin
            return (round(float(origin.X), 6), round(float(origin.Y), 6), round(float(origin.Z), 6))
        except Exception:
            return (999999.0, 999999.0, 999999.0)

    def ref_sort_key(ref_conn):
        try:
            return int(get_element_id_value(ref_conn.Owner))
        except Exception:
            return 999999999

    neighbor_ids = set()
    connectors = sorted(duct_wrap.get_connectors(), key=connector_sort_key)

    for connector in connectors:
        try:
            if not connector or not connector.IsConnected:
                continue
        except Exception:
            continue

        try:
            all_refs = sorted(list(connector.AllRefs), key=ref_sort_key)
        except Exception:
            continue

        for ref_conn in all_refs:
            try:
                connected_elem = ref_conn.Owner
            except Exception:
                continue

            neighbor_id = get_element_id_value(connected_elem)
            if neighbor_id == duct_wrap.id:
                continue
            if neighbor_id in straight_ids:
                neighbor_ids.add(neighbor_id)

    return neighbor_ids


def is_vertical_piece(elem):
    try:
        angle = RevitXYZ(elem).straight_joint_degree()
    except Exception as ex:
        non_fatal_warnings.append((elem, "Angle check failed: {}".format(ex)))
        return False

    return angle is not None and abs(angle) >= 85


def get_short_threshold_inches(duct_wrap):
    fam = (duct_wrap.family or '').strip().lower()
    conn0 = (duct_wrap.connector_0_type or '').strip().lower()
    conn1 = (duct_wrap.connector_1_type or '').strip().lower()

    if conn0 and conn1 and conn0 == conn1:
        for (k_family, k_conn), threshold in CONNECTOR_THRESHOLDS.items():
            if (
                fam == (k_family or '').strip().lower()
                and conn0 == (k_conn or '').strip().lower()
            ):
                return float(threshold)

    return float(DEFAULT_SHORT_THRESHOLD_IN)


def is_spiral_under_no_tag_length(duct_wrap):
    family_name = (duct_wrap.family or '').strip().lower()
    if 'spiral' not in family_name:
        return False

    length_in = duct_wrap.length
    if not isinstance(length_in, (int, float)):
        return False

    return float(length_in) < SPIRAL_NO_TAG_MAX_LENGTH_IN


def iter_components(node_ids, adjacency_map):
    remaining = set(node_ids)
    while remaining:
        start_id = min(remaining)
        stack = [start_id]
        component = []
        remaining.remove(start_id)

        while stack:
            current_id = stack.pop()
            component.append(current_id)
            for neighbor_id in sorted(adjacency_map.get(current_id, set()), reverse=True):
                if neighbor_id in remaining:
                    remaining.remove(neighbor_id)
                    stack.append(neighbor_id)

        yield component


def order_run_segment(component_ids, adjacency_map):
    component_ids = sorted(component_ids)
    component_set = set(component_ids)
    endpoints = []

    for elem_id in component_ids:
        degree = len(
            [neighbor_id for neighbor_id in adjacency_map.get(elem_id, set())
             if neighbor_id in component_set]
        )
        if degree <= 1:
            endpoints.append(elem_id)
        elif degree > 2:
            raise RuntimeError(
                "Non-linear run at element {}".format(elem_id)
            )

    endpoints.sort()

    ordered_ids = []
    visited = set()
    previous_id = None
    current_id = endpoints[0] if endpoints else component_ids[0]

    while current_id is not None:
        ordered_ids.append(current_id)
        visited.add(current_id)
        next_ids = sorted([
            neighbor_id for neighbor_id in adjacency_map.get(current_id, set())
            if neighbor_id in component_set and neighbor_id != previous_id
            and neighbor_id not in visited
        ])

        if len(next_ids) > 1:
            raise RuntimeError(
                "Non-linear run at element {}".format(current_id)
            )

        previous_id, current_id = (
            current_id,
            next_ids[0] if next_ids else None,
        )

    if len(ordered_ids) != len(component_ids):
        raise RuntimeError(
            "Could not order straight run containing element {}".format(
                component_ids[0]
            )
        )

    return ordered_ids


def add_slot_assignment(assignments, elem_id, slot_names):
    elem_slots = assignments.setdefault(elem_id, [])
    for slot_name in slot_names:
        if slot_name not in elem_slots:
            elem_slots.append(slot_name)


def should_singleton_use_all_three(
    elem_id,
    straight_wraps_by_id,
    straight_ids,
    split_boundary_ids,
    taggable_run_ids,
):
    duct_wrap = straight_wraps_by_id[elem_id]
    connectors = duct_wrap.get_connectors()
    if not connectors:
        return False

    boundary_kinds = []
    for connector_idx in range(len(connectors)):
        connected_elems = duct_wrap.get_connected_elements(connector_idx)
        if not connected_elems:
            boundary_kinds.append("open")
            continue

        has_split_straight = False
        has_taggable_straight = False
        has_other_straight = False
        has_non_straight = False

        for conn_elem in connected_elems:
            conn_id = get_element_id_value(conn_elem)
            if conn_id == elem_id:
                continue

            if conn_id in straight_ids:
                if conn_id in split_boundary_ids:
                    has_split_straight = True
                elif conn_id in taggable_run_ids:
                    has_taggable_straight = True
                else:
                    has_other_straight = True
            else:
                has_non_straight = True

        if has_split_straight:
            boundary_kinds.append("split")
        elif has_taggable_straight or has_other_straight:
            boundary_kinds.append("straight")
        elif has_non_straight:
            boundary_kinds.append("fitting")
        else:
            boundary_kinds.append("open")

    if any(kind in ("split", "straight") for kind in boundary_kinds):
        return False

    true_boundaries = [
        kind for kind in boundary_kinds if kind in ("open", "fitting")
    ]
    return len(true_boundaries) >= 2


def assign_slots_for_segment(
    ordered_ids,
    straight_wraps_by_id,
    straight_ids,
    split_boundary_ids,
    taggable_run_ids,
):
    assignments = {}
    segment_count = len(ordered_ids)

    if segment_count == 0:
        return assignments

    if segment_count == 1:
        singleton_id = ordered_ids[0]
        if should_singleton_use_all_three(
            singleton_id,
            straight_wraps_by_id,
            straight_ids,
            split_boundary_ids,
            taggable_run_ids,
        ):
            add_slot_assignment(assignments, singleton_id, full_run_tag_slots)
        else:
            add_slot_assignment(assignments, singleton_id,
                                [SLOT_LENGTH, SLOT_SIZE])
        return assignments

    if segment_count == 2:
        first_id, second_id = ordered_ids
        first_length = straight_wraps_by_id[first_id].length
        second_length = straight_wraps_by_id[second_id].length
        shorter_id = first_id
        longer_id = second_id

        if isinstance(first_length, (int, float)) and isinstance(second_length, (int, float)):
            if second_length < first_length:
                shorter_id = second_id
                longer_id = first_id
            elif second_length == first_length:
                shorter_id = min(first_id, second_id)
                longer_id = max(first_id, second_id)

        add_slot_assignment(assignments, shorter_id, [SLOT_LENGTH, SLOT_SIZE])
        add_slot_assignment(assignments, longer_id, [SLOT_BOD, SLOT_LENGTH])
        return assignments

    add_slot_assignment(assignments, ordered_ids[0], [SLOT_LENGTH, SLOT_SIZE])
    add_slot_assignment(assignments, ordered_ids[-1], [SLOT_LENGTH, SLOT_SIZE])
    add_slot_assignment(assignments, ordered_ids[1], [SLOT_BOD, SLOT_LENGTH])
    add_slot_assignment(assignments, ordered_ids[-2], [SLOT_BOD, SLOT_LENGTH])
    return assignments


def get_direction_vector(elem):
    loc = elem.Location
    if loc and hasattr(loc, 'Curve') and loc.Curve:
        curve = loc.Curve
        return (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize()
    return None


def get_offset_point(base_point, elem, tag_index, tag_spacing, warning_label):
    offset_pt = base_point
    try:
        dir_vec = get_direction_vector(elem)
        if dir_vec is not None:
            offset_distance = tag_index * tag_spacing
            offset_pt = XYZ(
                base_point.X + dir_vec.X * offset_distance,
                base_point.Y + dir_vec.Y * offset_distance,
                base_point.Z + dir_vec.Z * offset_distance,
            )
    except Exception as ex:
        non_fatal_warnings.append(
            (elem, "{}: {}".format(warning_label, ex))
        )
    return offset_pt


def rotate_tag_to_direction(elem, new_tag, offset_pt):
    try:
        dir_vec = get_direction_vector(elem)
        if dir_vec is None:
            return

        angle = math.atan2(dir_vec.Y, dir_vec.X)
        axis = Line.CreateBound(
            offset_pt,
            XYZ(offset_pt.X, offset_pt.Y, offset_pt.Z + 1),
        )
        ElementTransformUtils.RotateElement(doc, new_tag.Id, axis, angle)
    except Exception as ex:
        non_fatal_warnings.append((elem, "Rotation skipped: {}".format(ex)))


def place_tags_for_slots(elem, slot_names, existing_tag_fams, resolved_slot_specs):
    tagged_this_element = False
    any_tag_failed = False
    last_error = None
    tag_index = 0
    tag_spacing = 1.0

    for slot_name in slot_names:
        tag_spec = resolved_slot_specs.get(slot_name)
        if tag_spec is None:
            any_tag_failed = True
            last_error = "Unresolved slot [{}]".format(slot_name)
            continue

        tag_symbol, fam_name, matched_name = tag_spec

        try:
            if fam_name in existing_tag_fams:
                continue

            face_ref, face_pt = tagger.get_face_facing_view(elem)
            if face_ref is not None and face_pt is not None:
                offset_pt = get_offset_point(
                    face_pt,
                    elem,
                    tag_index,
                    tag_spacing,
                    "Face offset fallback used",
                )
                new_tag = tagger.place_tag(face_ref, tag_symbol, offset_pt)
            else:
                bbox = elem.get_BoundingBox(view)
                if bbox is None:
                    any_tag_failed = True
                    last_error = "No valid placement location found"
                    continue

                center = (bbox.Min + bbox.Max) / 2.0
                offset_pt = get_offset_point(
                    center,
                    elem,
                    tag_index,
                    tag_spacing,
                    "BBox offset fallback used",
                )
                new_tag = tagger.place_tag(elem, tag_symbol, offset_pt)

            tagged_this_element = True
            existing_tag_fams.add(fam_name)
            tag_index += 1
            rotate_tag_to_direction(elem, new_tag, offset_pt)
        except Exception as ex:
            any_tag_failed = True
            last_error = "{} [{}]".format(str(ex), matched_name)

    return tagged_this_element, any_tag_failed, last_error


def describe_slots(slot_names):
    return ", ".join(slot_names) if slot_names else "(none)"


t = Transaction(doc, "Tag Full Joints")
t.Start()
try:
    # Step 1: Get all ducts in view
    all_ducts_in_view = ducts
    output.print_md(
        "**Found {} total ducts in view**".format(len(all_ducts_in_view)))

    # Step 2: Build the straight-duct population used for run analysis.
    straight_ducts = [
        d for d in all_ducts_in_view if is_straight_family(d.family)
    ]
    output.print_md(
        "**Found {} straight ducts in view**".format(len(straight_ducts)))

    straight_wraps_by_id = {
        get_element_id_value(d.element): d for d in straight_ducts
    }
    straight_ids = set(straight_wraps_by_id.keys())
    straight_adjacency = {}
    short_length_only_ids = set()
    short_as_full_ids = set()
    spiral_no_tag_ids = set()
    vertical_ids = set()

    for elem_id, duct_wrap in straight_wraps_by_id.items():
        straight_adjacency[elem_id] = get_straight_neighbor_ids(
            duct_wrap,
            straight_ids,
        )

    # Normalize adjacency to undirected links so component detection is stable.
    for elem_id, neighbors in list(straight_adjacency.items()):
        for neighbor_id in list(neighbors):
            if neighbor_id == elem_id or neighbor_id not in straight_ids:
                continue
            straight_adjacency.setdefault(neighbor_id, set()).add(elem_id)

    for elem_id in list(straight_adjacency.keys()):
        straight_adjacency[elem_id].discard(elem_id)

    for elem_id, duct_wrap in straight_wraps_by_id.items():

        if is_vertical_piece(duct_wrap.element):
            vertical_ids.add(elem_id)
            skipped_vertical.append(duct_wrap)
            continue

        if is_spiral_under_no_tag_length(duct_wrap):
            spiral_no_tag_ids.add(elem_id)
            continue

        if duct_wrap.joint_size == JointSize.SHORT:
            threshold_in = get_short_threshold_inches(duct_wrap)
            quarter_threshold_in = threshold_in * SHORT_LENGTH_ONLY_RATIO
            length_in = duct_wrap.length

            if (
                isinstance(length_in, (int, float))
                and length_in < quarter_threshold_in
            ):
                short_length_only_ids.add(elem_id)
            else:
                short_as_full_ids.add(elem_id)

    output.print_md(
        "**Found {} short straights tagged as length-only (<25%)**".format(
            len(short_length_only_ids)
        ))
    output.print_md(
        "**Found {} short straights treated as full pieces (>=25%)**".format(
            len(short_as_full_ids)
        ))
    output.print_md(
        "**Found {} spiral straights with length under 12\" (no tags)**".format(
            len(spiral_no_tag_ids)
        ))

    taggable_run_ids = (
        straight_ids
        - vertical_ids
        - short_length_only_ids
        - spiral_no_tag_ids
    )
    split_boundary_ids = set(vertical_ids) | set(
        short_length_only_ids) | set(spiral_no_tag_ids)
    run_adjacency = {
        elem_id: set(
            neighbor_id for neighbor_id in straight_adjacency.get(elem_id, set())
            if neighbor_id in taggable_run_ids
        )
        for elem_id in taggable_run_ids
    }

    slot_assignments = {}
    for short_id in short_length_only_ids:
        add_slot_assignment(slot_assignments, short_id, [SLOT_LENGTH])

    ordered_run_segments = []
    for component_ids in iter_components(taggable_run_ids, run_adjacency):
        try:
            ordered_ids = order_run_segment(component_ids, run_adjacency)
        except RuntimeError as ex:
            non_fatal_warnings.append(
                (straight_wraps_by_id[component_ids[0]].element, str(ex))
            )
            continue

        ordered_run_segments.append(ordered_ids)
        segment_assignments = assign_slots_for_segment(
            ordered_ids,
            straight_wraps_by_id,
            straight_ids,
            split_boundary_ids,
            taggable_run_ids,
        )
        for elem_id, slot_names in segment_assignments.items():
            add_slot_assignment(slot_assignments, elem_id, slot_names)

    output.print_md(
        "**Found {} straight run segments to evaluate**".format(
            len(ordered_run_segments)
        ))

    no_tag_needed_ids = sorted(
        spiral_no_tag_ids
        | set(
            elem_id for elem_id in taggable_run_ids
            if elem_id not in slot_assignments
        )
    )
    for elem_id in no_tag_needed_ids:
        no_tag_needed.append(straight_wraps_by_id[elem_id])

    target_elems = [
        straight_wraps_by_id[elem_id].element
        for elem_id in sorted(slot_assignments.keys())
    ]
    existing_tag_map = tagger.build_existing_tag_family_map(target_elems)

    resolved_slot_specs = {}
    required_slots = []
    for slot_names in slot_assignments.values():
        for slot_name in slot_names:
            if slot_name not in required_slots:
                required_slots.append(slot_name)

    for tag_key in required_slots:
        tag_symbol, fam_name, matched_name = resolve_tag_slot(tag_key)
        if tag_symbol is None or not fam_name:
            continue
        resolved_slot_specs[tag_key] = (tag_symbol, fam_name, matched_name)

    if PRINT_RUN_ASSIGNMENT_DEBUG:
        output.print_md("## Run Assignment Debug")
        for elem_id in sorted(slot_assignments.keys()):
            wrap = straight_wraps_by_id[elem_id]
            output.print_md(
                "### ID: {} | Len: {:06.2f} | Fam: {} | Slots: {}".format(
                    output.linkify(wrap.element.Id),
                    wrap.length if isinstance(
                        wrap.length, (int, float)) else 0.0,
                    wrap.family,
                    describe_slots(slot_assignments.get(elem_id, [])),
                )
            )

        if no_tag_needed_ids:
            output.print_md("### No-Tag IDs: {}".format(
                output.linkify(
                    [straight_wraps_by_id[i].element.Id for i in no_tag_needed_ids])
            ))
        output.print_md("---")

    if missing_tag_labels:
        output.print_md(
            "# Warning: unresolved tag labels: {}".format(
                ", ".join(sorted(missing_tag_labels))
            )
        )

    if not resolved_slot_specs and slot_assignments:
        output.print_md(
            "Tag placement error: no tag labels could be resolved from current configuration")
        t.RollBack()
        raise RuntimeError("No tag labels resolved")

    # Step 4: Place only the slots assigned to each straight piece by run rules.
    for elem_id in sorted(slot_assignments.keys()):
        elem_wrap = straight_wraps_by_id[elem_id]
        elem = elem_wrap.element
        existing_tag_fams = set(existing_tag_map.get(elem_id, set()))
        slot_names = slot_assignments.get(elem_id, [])

        tagged_this_element, any_tag_failed, last_error = place_tags_for_slots(
            elem,
            slot_names,
            existing_tag_fams,
            resolved_slot_specs,
        )

        # Add element to appropriate list
        if tagged_this_element:
            needs_tagging.append(elem_wrap)
            if any_tag_failed and last_error:
                non_fatal_warnings.append(
                    (elem, "Partially tagged: {}".format(last_error)))
        elif any_tag_failed and last_error:
            failed_to_tag.append((elem, last_error))
        else:
            already_tagged.append(elem_wrap)
    output.print_md("---")

    if skipped_vertical:
        output.print_md("## Skipped Vertical")
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

    if no_tag_needed:
        output.print_md("## No Tag Needed")
        for i, d in enumerate(no_tag_needed, start=1):
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
    if skipped_vertical:
        skipped_ids = [d.element.Id for d in skipped_vertical]
        output.print_md("# Skipped vertical: {}, {}".format(
            len(skipped_vertical), output.linkify(skipped_ids)))
    if no_tag_needed:
        no_tag_ids = [d.element.Id for d in no_tag_needed]
        output.print_md("# No tag needed: {}, {}".format(
            len(no_tag_needed), output.linkify(no_tag_ids)))

    if non_fatal_warnings:
        output.print_md(
            "# Non-fatal warnings: {}".format(len(non_fatal_warnings)))

    all_ducts = needs_tagging + already_tagged + skipped_vertical + no_tag_needed
    all_ids = [d.element.Id for d in all_ducts]
    output.print_md("# Total: {}, {}".format(
        len(all_ducts), output.linkify(all_ids)))
    t.Commit()
except Exception as e:
    output.print_md("Tag placement error: {}".format(e))
    t.RollBack()
    raise
