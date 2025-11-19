# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from pyrevit import revit, script, forms
from revit_duct import RevitDuct
from revit_parameter import RevitParameter
from Autodesk.Revit.DB import Transaction

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
ducts = RevitDuct.from_selection(uidoc, doc, view)
rp = RevitParameter(doc, app)


# Code
# ==================================================
if not ducts:
    forms.alert("Select fabrication ducts first.", exitscript=True)

family_list = ["transition", "mitred offset",
               "radius offset", "mitered offset",
               "ogee", "offset", ]

offset_list = ["ogee", "offset", "radius offset",
               "mitered offset", "mitred offset"]


with Transaction(doc, "Offset Parameter") as t:
    t.Start()

    for d in ducts:
        tag = None
        output.print_md("start of main loop")
        family = d.family.lower().strip()
        if family in family_list:
            output.print_md("start of family loop")
            cen_w = d.centerline_width
            cen_h = d.centerline_height
            off_l = d.offset_left
            off_r = d.offset_right
            off_t = d.offset_top
            off_b = d.offset_bottom
            off_d = d.offset_height
            h_in = d.heigth_in
            h_out = d.heigth_out
            top_e = d.top_edge_rise_in()
            bot_e = d.bottom_edge_rise_in()

            output.print_md("offset width: {}".format(d.offset_width))
            output.print_md("offset height: {}".format(d.offset_height))

            output.print_md("cen_w: {}, cen_h: {}".format(cen_w, cen_h))
            output.print_md("off_l: {}, off_r: {}, off_t: {}, off_b: {}".format(
                off_l, off_r, off_t, off_b))
            output.print_md("top_e: {}, bot_e: {}".format(top_e, bot_e))

            if family == "transition":
                output.print_md("family:{}".format(family))
                if abs(cen_h) < 0.01:
                    output.print_md("cen_h:{}".format(cen_h))
                    tag = "CL"
                elif abs(off_d) < 0.01:
                    output.print_md("off_d:{}".format(off_d))
                    tag = "FOB"
                elif h_in - off_d - h_out == 0:
                    output.print_md(
                        "he_in:{}, off_d:{}, he_out:{}".format(h_in, off_d, h_out))
                    tag = "FOT"
                elif off_d + h_out > h_in:
                    dif = off_d + h_out - h_in
                    output.print_md(
                        "off_d:{}↑".format(dif))
                    tag = "{}↑".format(dif)
                elif off_d + h_out > h_in:
                    output.print_md(
                        "off_d:{}↓".format(off_d))
                    tag = "{}↓".format(off_b)

            elif family in offset_list:
                output.print_md("family:{}, cen_w:{}".format(family, cen_w))
                tag = "{}→".format(cen_w)

        if tag is not None:
            rp.set_parameter_value(d.element, "_jfn_offset", tag)
            output.print_md("{}".format(tag))
    t.Commit()
