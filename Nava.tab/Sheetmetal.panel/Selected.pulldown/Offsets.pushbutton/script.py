# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from revit_duct import RevitDuct
from revit_output import print_parameter_help
from pyrevit import revit, script, forms

# Button display information
# =================================================
__title__ = "Offset info"
__doc__ = """
Gives offset information about specific selected duct.
"""

# Variables
# ==================================================
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
view = revit.active_view
output = script.get_output()

# Code
# ==================================================
ducts = RevitDuct.from_selection(uidoc, doc, view)

family_list = {
    "transition",
    "mitred offset",
    "radius offset",
    "mitered offset",
    "ogee",
    "offset"
}

if not ducts:
    forms.alert("Please select one or more ducts first.")
    raise SystemExit

# Normalize family names once (lower + strip) and filter
filtered_ducts = [d for d in ducts if (
    d.family or "").lower().strip() in family_list]

for d in filtered_ducts:
    data = d.offset_data
    link = output.linkify(d.element.Id)

    if not data:
        output.print_md(
            "\n**Duct:** {} — no offset data available.".format(link))
        continue

    edges = data.get('edges') or {}
    is_round = bool(edges.get('round'))
    output.print_md("\n**Duct:** {}".format(link))

    if is_round:
        diam_in = edges.get('diam_in') or d.diameter_in
        diam_out = edges.get('diam_out') or d.diameter_out or diam_in
        if diam_in and diam_out:
            output.print_md(
                "Diameter: {0:.2f}\" → {1:.2f}\"".format(diam_in, diam_out))
        else:
            output.print_md("Diameter: (unknown)")
    else:
        w_in = d.width_in or 0.0
        h_in = d.heigth_in or 0.0
        w_out = d.width_out or w_in
        h_out = d.heigth_out or h_in
        output.print_md("Size: {0:.2f}x{1:.2f} → {2:.2f}x{3:.2f}".format(
            w_in, h_in, w_out, h_out))

    # Centerline offsets taken from dict
    cw = data.get('centerline_width') or 0.0
    ch = data.get('centerline_height') or 0.0
    output.print_md("\n**Centerline Offsets:**")
    output.print_md("Plan: {0:.2f}\" | Vertical: {1:.2f}\"".format(cw, ch))

    if not is_round:
        whole = edges.get('whole_in') or {}
        output.print_md("\n**Edge Offsets (Whole Inches):**")

        def fmt_edge(val):
            return str(val) if val is not None else "N/A"
        output.print_md(
            "Left: {0}\" | Right: {1}\" | Top: {2}\" | Bottom: {3}\"".format(
                fmt_edge(
                    whole.get('left')), fmt_edge(
                    whole.get('right')), fmt_edge(
                    whole.get('top')), fmt_edge(
                        whole.get('bottom'))))
    else:
        output.print_md("\n_No rectangular edge offsets (round)._")

# Final print statements
print_parameter_help(output)
