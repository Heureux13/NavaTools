# # -*- coding: utf-8 -*-
# # ======================================================================
# """Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

# This code and associated documentation files may not be copied, modified,
# distributed, or used in any form without the prior written permission of
# the copyright holder."""
# # ======================================================================

# # Imports
# # ==================================================
# from revit_tagging import RevitTagging
# from Autodesk.Revit.DB import Transaction, ElementTransformUtils, XYZ, Line
# from revit_duct import RevitDuct
# from revit_xyz import RevitXYZ
# from pyrevit import revit, script
# import math
# from size import Size

# # Button display information
# # =================================================
# __title__ = "Tag Size"
# __doc__ = """
# Adds size tags to ducts
# """
# # Universal Variables
# # ====================================================
# uidoc = __revit__.ActiveUIDocument
# app = __revit__.Application

# doc = revit.doc
# view = revit.active_view
# output = script.get_output()
# ducts = RevitDuct.all(doc, view)
# tagger = RevitTagging(doc, view)


# # Helpers
# # ===================================================
# def _int_id(el):
#     """Return an int ElementId value across Revit versions (IntegerValue/Value)."""
#     try:
#         eid = getattr(el, 'Id', None)
#         if eid is None:
#             return None
#         iv = getattr(eid, 'IntegerValue', None)
#         if iv is not None:
#             return iv
#         val = getattr(eid, 'Value', None)
#         return val
#     except Exception:
#         return None


# def trace_to_spiral(elem):
#     """Iteratively trace connections to find spiral duct elements."""
#     visited = set()
#     queue = [elem]
#     spirals = []
#     max_iterations = 100  # Safety limit
#     iterations = 0

#     while queue and iterations < max_iterations:
#         iterations += 1
#         current = queue.pop(0)
#         current_id = _int_id(current)

#         if current_id is None or current_id in visited:
#             continue
#         visited.add(current_id)

#         try:
#             rd_elem = RevitDuct(doc, view, current)
#             fam = (rd_elem.family or '').strip().lower()

#             # If this is a spiral duct, collect it
#             if fam in round_straight_families:
#                 spirals.append(current)
#                 continue

#             # If gored elbow, add connected elements to queue
#             if fam == 'gored elbow':
#                 connectors = rd_elem.get_connectors() or []
#                 for connector in connectors:
#                     try:
#                         if not connector or not hasattr(connector, 'AllRefs'):
#                             continue
#                         refs = connector.AllRefs
#                         if not refs:
#                             continue
#                         for ref_conn in refs:
#                             if not ref_conn or not hasattr(ref_conn, 'Owner'):
#                                 continue
#                             owner = ref_conn.Owner
#                             if not owner:
#                                 continue
#                             owner_id = _int_id(owner)
#                             if owner_id is not None and owner_id != current_id and owner_id not in visited:
#                                 queue.append(owner)
#                     except Exception:
#                         continue
#         except Exception:
#             continue

#     return spirals


# # List / Dics
# # ========================================================================
# square_straight_families = {
#     "straight",
# }

# square_fitting_families = {
#     'transition',
#     'square to ø',
#     'drop cheek',
#     'ogee',
#     'offset',
#     'end cap',
#     'tdf end cap'
# }

# square_check = {
#     'family': {'straight', 'straight duct'},
# }

# round_straight_families = {
#     'spiral tube',
#     'round duct',
#     'spiral duct',
# }

# round_fitting_families = {
#     'boot tap',
#     'boot saddle tap',
#     'boot tap - wdamper',
#     'conical tap',
#     'conical saddle tap',
#     'conical tap - wdamper',
#     'reducer',
#     'square to ø',
#     '45 tap',
# }

# round_check = {
#     'family': {'spiral duct', 'gored elbow'}
# }

# size_tags = {
#     "_umi_size",
# }

# # Code
# # ==================================================


# round_fittings = [
#     d for d in ducts
#     if d.family and d.family.strip().lower() in round_fitting_families
# ]

# square_fittings = [
#     d for d in ducts
#     if d.family and d.family.strip().lower() in square_fitting_families
# ]

# tagging_list = []

# for d in round_fittings:
#     try:
#         connectors = d.get_connectors() or []
#         for i in range(len(connectors)):
#             temp_connected = []
#             filtered_connected = []

#             try:
#                 for connected in (d.get_connected_elements(connector_index=i) or []):
#                     try:
#                         rd = RevitDuct(doc, view, connected)
#                         if (rd.category or '').strip().lower() == 'mep fabrication ductwork':
#                             temp_connected.append(connected)
#                     except Exception:
#                         continue
#             except Exception:
#                 continue

#             allowed = round_check.get('family', set())

#             for duct in temp_connected:
#                 try:
#                     rd2 = RevitDuct(doc, view, duct)
#                     fam2 = (rd2.family or '').strip().lower()
#                     if fam2 in allowed:
#                         filtered_connected.append(duct)
#                 except Exception:
#                     continue

#             spirals_with_size = []
#             for elem in filtered_connected:
#                 try:
#                     for spiral_elem in trace_to_spiral(elem):
#                         spirals_with_size.append(spiral_elem)
#                 except Exception:
#                     continue

#             # Build a stable, unique list of spiral straights to tag
#             uniq_spirals = {}
#             for se in spirals_with_size:
#                 sid = _int_id(se)
#                 if sid is not None and sid not in uniq_spirals:
#                     uniq_spirals[sid] = se
#             for sid in sorted(uniq_spirals.keys()):
#                 tagging_list.append(uniq_spirals[sid])
#     except Exception:
#         continue


# for d in square_fittings:
#     try:
#         connectors = d.get_connectors() or []
#         for i in range(len(connectors)):
#             temp_connected = []
#             filtered_connected = []

#             try:
#                 for connected in (d.get_connected_elements(connector_index=i) or []):
#                     try:
#                         sqr = RevitDuct(doc, view, connected)
#                         if (sqr.category or '').strip().lower() == 'mep fabrication ductwork':
#                             temp_connected.append(connected)
#                     except Exception:
#                         continue
#             except Exception:
#                 continue

#             allowed = square_check.get('family', set())

#             for duct in temp_connected:
#                 try:
#                     sqr2 = RevitDuct(doc, view, duct)
#                     fam2 = (sqr2.family or '').strip().lower()
#                     if fam2 in allowed:
#                         filtered_connected.append(duct)
#                 except Exception:
#                     continue

#             # Determine square straights connected to this fitting using Size shape detection
#             square_connected = []
#             for elem in filtered_connected:
#                 try:
#                     rd_conn = RevitDuct(doc, view, elem)
#                     sz_conn = Size(rd_conn.size)
#                     conn_shape = sz_conn.in_shape() or sz_conn.out_shape()
#                     if conn_shape == 'rectangle':
#                         square_connected.append(elem)
#                 except Exception:
#                     continue

#             try:
#                 fam_fit = (d.family or '').strip().lower()
#                 sz_fit = Size(d.size)
#                 in_shape = sz_fit.in_shape()
#                 out_shape = sz_fit.out_shape()

#                 # Special case: "square to Ø" — only tag the square side
#                 if fam_fit == 'square to ø':
#                     if square_connected:
#                         tagging_list.append(square_connected[0])
#                     continue

#                 # For other square fittings, tag at least two pieces when both sides are square
#                 if in_shape == 'rectangle' and out_shape == 'rectangle':
#                     if len(square_connected) >= 2:
#                         tagging_list.append(square_connected[0])
#                         tagging_list.append(square_connected[1])
#                     elif len(square_connected) == 1:
#                         tagging_list.append(square_connected[0])
#             except Exception:
#                 continue
#     except Exception:
#         continue

# # Stabilize and pre-filter tagging list to ensure idempotence across runs
# # 1) Deduplicate by ElementId
# stable_map = {}
# for el in tagging_list:
#     eid = _int_id(el)
#     if eid is not None and eid not in stable_map:
#         stable_map[eid] = el
# tagging_list = [stable_map[eid] for eid in sorted(stable_map.keys())]

# # 2) Exclude elements already tagged with any of the size tag families
# try:
#     tag_syms_map = {}
#     for name in size_tags:
#         try:
#             tag_syms_map[name] = tagger.get_label(name)
#         except Exception:
#             continue
#     _filtered = []
#     for el in tagging_list:
#         skip = False
#         for name, sym in tag_syms_map.items():
#             famname = getattr(getattr(sym, 'Family', None), 'Name', '')
#             if famname and tagger.already_tagged(el, famname):
#                 skip = True
#                 break
#         if not skip:
#             _filtered.append(el)
#     tagging_list = _filtered
# except Exception:
#     # if any lookup fails, proceed without pre-filtering
#     pass

# # Tagging phase: place size tags on all collected elements
# t = None
# try:
#     t = Transaction(doc, "Tag Size")
#     t.Start()

#     for elem in tagging_list:
#         for tag_name in size_tags:
#             try:
#                 tag_sym = tagger.get_label(tag_name)
#                 # avoid duplicates of the same tag family on an element
#                 if tagger.already_tagged(elem, getattr(tag_sym.Family, 'Name', '')):
#                     continue

#                 # choose a good reference/point; prefer a face facing the view
#                 ref, pt = tagger.get_face_facing_view(elem)
#                 if pt is None:
#                     # fallback: use midpoint along element curve, or bbox center
#                     try:
#                         rd = RevitDuct(doc, view, elem)
#                         pt = RevitTagging.midpoint_location(rd, 0.5, 0.0)
#                     except Exception:
#                         pass
#                 if pt is None:
#                     try:
#                         bbox = elem.get_BoundingBox(view)
#                         if bbox:
#                             center = (bbox.Min + bbox.Max) / 2.0
#                             pt = XYZ(center.X, center.Y, center.Z)
#                     except Exception:
#                         pass
#                 if pt is None:
#                     continue

#                 anchor = ref if ref is not None else elem
#                 tagger.place_tag(anchor, tag_symbol=tag_sym, point_xyz=pt)
#             except Exception:
#                 # skip problematic element/tag gracefully
#                 continue

#     t.Commit()
# except Exception:
#     if t and t.HasStarted() and not t.HasEnded():
#         try:
#             t.RollBack()
#         except Exception:
#             pass
