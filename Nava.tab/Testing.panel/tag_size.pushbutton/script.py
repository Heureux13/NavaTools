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
from size import Size

# Button display information
# =================================================
__title__ = "Tag Size"
__doc__ = """
Adds size tags to ducts
"""
# Universal Variables
# ====================================================
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

doc = revit.doc
view = revit.active_view
output = script.get_output()
ducts = RevitDuct.all(doc, view)
tagger = RevitTagging(doc, view)


# Helpers
# ===================================================
def _int_id(el):
    """Return an int ElementId value across Revit versions (IntegerValue/Value)."""
    try:
        eid = getattr(el, 'Id', None)
        if eid is None:
            return None
        iv = getattr(eid, 'IntegerValue', None)
        if iv is not None:
            return iv
        val = getattr(eid, 'Value', None)
        return val
    except Exception:
        return None


def trace_to_spiral(elem, visited=None):
    if visited is None:
        visited = set()

    elem_id = _int_id(elem)
    if elem_id in visited:
        return []
    visited.add(elem_id)

    rd_elem = RevitDuct(doc, view, elem)
    fam = (rd_elem.family or '').strip().lower()

    # If this is a spiral duct, return it
    if fam in round_straight_families:
        return [elem]

    # If gored elbow, trace its connections
    if fam == 'gored elbow':
        spirals = []
        connectors = rd_elem.get_connectors() or []
        for connector in connectors:
            for ref_conn in connector.AllRefs:
                owner_id = _int_id(ref_conn.Owner)
                if owner_id != elem_id:
                    spirals.extend(trace_to_spiral(
                        ref_conn.Owner, visited))
        return spirals

    return []


# List / Dics
# ========================================================================
square_straight_families = {
    "straight",
}

square_fitting_families = {
    'transition',
    'square to ø',
    'drop cheek',
    'ogee',
    'offset',
    'end cap',
    'tdf end cap'
}

square_check = {
    'family': {'straight', 'straight duct'},
}

round_straight_families = {
    'spiral tube',
    'round duct',
    'spiral duct',
}

round_fitting_families = {
    'boot tap',
    'boot saddle tap',
    'boot tap - wdamper',
    'conical tap',
    'conical saddle tap',
    'conical tap - wdamper',
    'reducer',
    'square to ø',
    '45 tap',
}

round_check = {
    'family': {'spiral duct', 'gored elbow'}
}

size_tags = {
    "_umi_size",
}

# Code
# ==================================================


round_fittings = [
    d for d in ducts
    if d.family and d.family.strip().lower() in round_fitting_families
]

square_fittings = [
    d for d in ducts
    if d.family and d.family.strip().lower() in square_fitting_families
]

tagging_list = []

for d in round_fittings:
    connectors = d.get_connectors() or []
    for i in range(len(connectors)):
        temp_connected = []
        filtered_connected = []

        for connected in (d.get_connected_elements(connector_index=i) or []):
            rd = RevitDuct(doc, view, connected)
            if (rd.category or '').strip().lower() == 'mep fabrication ductwork':
                temp_connected.append(connected)

        allowed = round_check.get('family', set())

        for duct in temp_connected:
            rd2 = RevitDuct(doc, view, duct)
            fam2 = (rd2.family or '').strip().lower()
            if fam2 in allowed:
                filtered_connected.append(duct)

        spirals_with_size = []
        for elem in filtered_connected:
            for spiral_elem in trace_to_spiral(elem):
                spirals_with_size.append(spiral_elem)

        # Build a stable, unique list of spiral straights to tag
        uniq_spirals = {}
        for se in spirals_with_size:
            sid = _int_id(se)
            if sid is not None and sid not in uniq_spirals:
                uniq_spirals[sid] = se
        for sid in sorted(uniq_spirals.keys()):
            tagging_list.append(uniq_spirals[sid])


for d in square_fittings:
    connectors = d.get_connectors() or []
    for i in range(len(connectors)):
        temp_connected = []
        filtered_connected = []

        for connected in (d.get_connected_elements(connector_index=i) or []):
            sqr = RevitDuct(doc, view, connected)
            if (sqr.category or '').strip().lower() == 'mep fabrication ductwork':
                temp_connected.append(connected)

        allowed = square_check.get('family', set())

        for duct in temp_connected:
            sqr2 = RevitDuct(doc, view, duct)
            fam2 = (sqr2.family or '').strip().lower()
            if fam2 in allowed:
                filtered_connected.append(duct)
        # Determine square straights connected to this fitting using Size shape detection
        square_connected = []
        for elem in filtered_connected:
            rd_conn = RevitDuct(doc, view, elem)
            sz_conn = Size(rd_conn.size)
            conn_shape = sz_conn.in_shape() or sz_conn.out_shape()
            if conn_shape == 'rectangle':
                square_connected.append(elem)

        fam_fit = (d.family or '').strip().lower()
        sz_fit = Size(d.size)
        in_shape = sz_fit.in_shape()
        out_shape = sz_fit.out_shape()

        # Special case: "square to Ø" — only tag the square side
        if fam_fit == 'square to ø':
            if square_connected:
                tagging_list.append(square_connected[0])
            continue

        # For other square fittings, tag at least two pieces when both sides are square
        if in_shape == 'rectangle' and out_shape == 'rectangle':
            if len(square_connected) >= 2:
                tagging_list.append(square_connected[0])
                tagging_list.append(square_connected[1])
            elif len(square_connected) == 1:
                tagging_list.append(square_connected[0])

# Stabilize and pre-filter tagging list to ensure idempotence across runs
# 1) Deduplicate by ElementId
stable_map = {}
for el in tagging_list:
    eid = _int_id(el)
    if eid is not None and eid not in stable_map:
        stable_map[eid] = el
tagging_list = [stable_map[eid] for eid in sorted(stable_map.keys())]

# 2) Exclude elements already tagged with any of the size tag families
try:
    tag_syms_map = {}
    for name in size_tags:
        try:
            tag_syms_map[name] = tagger.get_label(name)
        except Exception:
            continue
    _filtered = []
    for el in tagging_list:
        skip = False
        for name, sym in tag_syms_map.items():
            famname = getattr(getattr(sym, 'Family', None), 'Name', '')
            if famname and tagger.already_tagged(el, famname):
                skip = True
                break
        if not skip:
            _filtered.append(el)
    tagging_list = _filtered
except Exception:
    # if any lookup fails, proceed without pre-filtering
    pass

# Tagging phase: place size tags on all collected elements
try:
    t = Transaction(doc, "Tag Size")
    t.Start()

    for elem in tagging_list:
        for tag_name in size_tags:
            try:
                tag_sym = tagger.get_label(tag_name)
                # avoid duplicates of the same tag family on an element
                if tagger.already_tagged(elem, getattr(tag_sym.Family, 'Name', '')):
                    continue

                # choose a good reference/point; prefer a face facing the view
                ref, pt = tagger.get_face_facing_view(elem)
                if pt is None:
                    # fallback: use midpoint along element curve, or bbox center
                    rd = RevitDuct(doc, view, elem)
                    pt = RevitTagging.midpoint_location(rd, 0.5, 0.0)
                if pt is None:
                    bbox = elem.get_BoundingBox(view)
                    if bbox:
                        center = (bbox.Min + bbox.Max) / 2.0
                        pt = XYZ(center.X, center.Y, center.Z)
                if pt is None:
                    continue

                anchor = ref if ref is not None else elem
                tagger.place_tag(anchor, tag_symbol=tag_sym, point_xyz=pt)
            except Exception:
                # skip problematic element/tag gracefully
                continue

    t.Commit()
except Exception:
    pass
