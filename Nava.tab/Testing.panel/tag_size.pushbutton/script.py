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
def trace_to_spiral(elem, visited=None):
    if visited is None:
        visited = set()

    elem_id = elem.Id.IntegerValue if hasattr(
        elem.Id, 'IntegerValue') else elem.Id.Value
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
                if ref_conn.Owner.Id.IntegerValue != elem_id:
                    spirals.extend(trace_to_spiral(
                        ref_conn.Owner, visited))
        return spirals

    return []


# List / Dics
# ========================================================================
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
    '45 tap'
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

        if len(spirals_with_size) == 1:
            tagging_list.append(spirals_with_size[0])

        elif len(spirals_with_size) >= 2:
            diameters = []
            for elem in spirals_with_size:
                rd_sp = RevitDuct(doc, view, elem)
                dia = rd_sp.diameter_in if rd_sp.diameter_in is not None else float(
                    'inf')
                diameters.append((dia, elem))

            min_dia = min(diameters, key=lambda t: t[0])[0]
            for dia, elem in diameters:
                if dia == min_dia:
                    tagging_list.append(elem)
