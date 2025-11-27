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
wtf
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

DEBUG = False


def dbg(msg, *args):
    if not DEBUG:
        return
    if args:
        safe_args = []
        for a in args:
            try:
                if isinstance(a, (int, float)):
                    safe_args.append(float(a))
                else:
                    safe_args.append(a)
            except Exception:
                safe_args.append(a)
        try:
            output.print_md(msg.format(*safe_args))
        except Exception:
            output.print_md("{} | {}".format(
                msg, ", ".join([str(x) for x in safe_args])))
    else:
        output.print_md(msg)


# Code
# ==================================================
if not ducts:
    forms.alert("Select fabrication ducts first.", exitscript=True)

family_list = ["transition", "mitred offset",
               "radius offset", "mitered offset",
               "ogee", "offset", "reducer"]

reducer_square = ["transition"]

reducer_round = ["reducer"]

offset_list = ["ogee", "offset", "radius offset",
               "mitered offset", "mitred offset"]


with Transaction(doc, "Offset Parameter") as t:
    t.Start()

    for d in ducts:
        tag = None
        family = d.family.lower().strip()

        if family in family_list:
            # Read needed parameters
            oge_o = d.ogee_offset

            # Get offset data from RevitDuct method
            offset_data = d.classify_offset()
            if not offset_data:
                continue

            cen_w = offset_data['centerline_w']
            cen_h = offset_data['centerline_h']
            top_e = offset_data['top_edge']
            bot_e = offset_data['bot_edge']
            off_t = offset_data['top_mag']
            off_b = offset_data['bot_mag']
            top_aligned = offset_data['top_aligned']
            bot_aligned = offset_data['bot_aligned']
            cl_vert = offset_data['cl_vert']

            # Optional debug
            dbg("cen_w: {:.2f} in, cen_h: {:.2f} in", cen_w, cen_h)
            dbg("top_e: {:+.2f}\" bot_e: {:+.2f}\"", top_e, bot_e)

            # Classification
            if family in reducer_square:
                # Check if this is a pure rotation (centerline flat, equal opposite edge movement)
                is_rotation = (cen_h < 0.5) and abs(
                    abs(top_e) - abs(bot_e)) < 0.5

                if cl_vert or is_rotation:
                    tag = "CL"
                elif bot_aligned:
                    tag = "FOB"
                elif top_aligned:
                    tag = "FOT"
                else:
                    # Arrow based on top edge movement
                    mag = int(round(abs(top_e)))
                    if mag == 0:
                        tag = "CL"
                    else:
                        tag = (u'↑{}"'.format(mag)) if top_e > 0 else (
                            u'↓{}"'.format(mag))

            elif family in reducer_round:
                y_off = d.reducer_offset
                d_in = d.diameter_in
                d_out = d.diameter_out
                # output.print_md("y_off: {}\nd_in: {}\nd_out: {}".format((y_off, d_in, d_out)))

                # Validate data
                if y_off is not None and d_in and d_out:
                    expected_cl = (d_in - d_out) / 2.0

                    if abs(y_off - expected_cl) < 0.01:
                        tag = "CL"
                    elif abs(d_out + y_off - d_in) < 0.01:
                        tag = "FOS"
                    elif abs(y_off) < 0.1:
                        tag = "FOS"
                    else:
                        tag = (u'{}"→'.format(int(round(y_off))))

            elif family in offset_list:
                # Horizontal offset family; use ogee offset parameter or calculated centerline
                offset = oge_o or cen_w or 0
                tag = '{}"→'.format(int(round(offset)))

        if tag is not None:
            rp.set_parameter_value(d.element, "_Offset", tag)

            def f(x):
                try:
                    return float(x)
                except BaseException:
                    return 0.0

            ch = f(cen_h)
            te = f(top_e)
            be = f(bot_e)
            ot = int(off_t or 0)
            ob = int(off_b or 0)
            try:
                eid = d.element.Id.IntegerValue
            except BaseException:
                eid = d.element.Id

            output.print_md(
                "{} | {} | CLh:{:.0f}\" | T:{}\" B:{}\" | ΔT:{:+.0f}\" ΔB:{:+.0f}\" | {}".format(
                    eid, family, ch, ot, ob, te, be, tag
                )
            )

    t.Commit()
