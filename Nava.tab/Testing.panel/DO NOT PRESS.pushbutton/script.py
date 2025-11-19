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

DEBUG = True


def dbg(msg, *args):
    if not DEBUG:
        return
    if args:
        safe_args = []
        for a in args:
            try:
                # Coerce numerics to float so {:.2f} works in IronPython
                if isinstance(a, (int, float)):
                    safe_args.append(float(a))
                else:
                    safe_args.append(a)
            except Exception:
                safe_args.append(a)
        try:
            output.print_md(msg.format(*safe_args))
        except Exception:
            # Fallback: dump raw values if format fails
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
               "ogee", "offset", ]

offset_list = ["ogee", "offset", "radius offset",
               "mitered offset", "mitred offset"]


with Transaction(doc, "Offset Parameter") as t:
    t.Start()

    for d in ducts:
        tag = None
        dbg("start of main loop")
        family = d.family.lower().strip()
        if family in family_list:
            dbg("start of family loop")
            cen_w = d.centerline_width
            cen_h = d.centerline_height
            off_l = d.offset_left
            off_r = d.offset_right
            off_t = d.offset_top
            off_b = d.offset_bottom
            off_d = d.offset_height
            off_w = d.offset_width
            h_in = d.heigth_in
            h_out = d.heigth_out
            oge_o = d.ogee_offset
            top_e = d.top_edge_rise_in()
            bot_e = d.bottom_edge_rise_in()

            output.print_md("offset width: {}".format(d.offset_width))
            output.print_md("offset height: {}".format(d.offset_height))
            output.print_md("cen_w: {}, cen_h: {}".format(cen_w, cen_h))
            output.print_md("off_l: {}, off_r: {}, off_t: {}, off_b: {}".format(
                off_l, off_r, off_t, off_b))
            output.print_md("top_e: {}, bot_e: {}".format(top_e, bot_e))
            output.print_md("check: off_t={}, off_b={}, top_e={}, bot_e={}, cen_h={}".format(
                off_t, off_b, top_e, bot_e, cen_h))

            # Compact/optional debug
            dbg("cen_w: {}, cen_h: {}", cen_w, cen_h)
            dbg("edges: L:{} R:{} T:{} B:{}", off_l, off_r, off_t, off_b)
            dbg("rise: ΔTop:{:+.2f}\" ΔBot:{:+.2f}\"", top_e or 0, bot_e or 0)

            # Optional inlet/outlet diagnostics
            c_in, c_out = d.identify_inlet_outlet()
            dbg("IN: z={:.2f} | OUT: z={:.2f}", c_in.Origin.Z, c_out.Origin.Z)
            dbg("C0: z={:.2f} | C1: z={:.2f}", d.get_connector(
                0).Origin.Z, d.get_connector(1).Origin.Z)
            dbg("PRIMARY: w={}, h={} | SECONDARY: w={}, h={}",
                h_in, h_out, h_out, h_in)

            # Deterministic inlet/outlet
            c_in, c_out = d.identify_inlet_outlet()
            if not c_in or not c_out:
                continue

            p_in = c_in.Origin
            p_out = c_out.Origin

            # Horizontal centerline offset (plan length between connector origins)
            dx = p_out.X - p_in.X
            dy = p_out.Y - p_in.Y
            cen_w = (dx*dx + dy*dy) ** 0.5 * 12.0   # inches

            # Vertical centerline offset magnitude (ignore sign for CL test)
            dz = p_out.Z - p_in.Z
            cen_h = abs(dz) * 12.0                  # inches

            # Sizes
            h_in = d.heigth_in or 0.0
            h_out = d.heigth_out or h_in

            # World Z planes (feet)
            top_in_z = p_in.Z + 0.5 * (h_in / 12.0)
            top_out_z = p_out.Z + 0.5 * (h_out / 12.0)
            bot_in_z = p_in.Z - 0.5 * (h_in / 12.0)
            bot_out_z = p_out.Z - 0.5 * (h_out / 12.0)

            # Edge rises (inches, signed: +up, -down)
            top_e = (top_out_z - top_in_z) * 12.0
            bot_e = (bot_out_z - bot_in_z) * 12.0

            # Classification tolerance (inches)
            tol_in = 0.01

            top_aligned = abs(top_e) < tol_in
            bot_aligned = abs(bot_e) < tol_in
            cl_vert = top_aligned and bot_aligned

            # Make whole-inch magnitudes for display (do NOT use for alignment logic)
            off_t = int(round(abs(top_e)))
            off_b = int(round(abs(bot_e)))

            # Optional debug
            dbg("cen_w: {:.2f} in, cen_h: {:.2f} in", cen_w, cen_h)
            dbg("top_e: {:+.2f}\" bot_e: {:+.2f}\"", top_e, bot_e)

            # Classification (quiet; only debug when needed)
            if family == "transition":
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
                        tag = (u"↑{}\"".format(mag)) if top_e > 0 else (
                            u"↓{}\"".format(mag))

            elif family == "offset":
                # Horizontal offset family; use horizontal magnitude
                offset = oge_o or cen_w or 0
                output.print_md(
                    "DELETE THIS: oge_o:{}, cen_w:{}".format(oge_o, cen_w))
                tag = "{}→".format(int(round(offset)))

        if tag is not None:
            rp.set_parameter_value(d.element, "_Offset", tag)

            def f(x):
                try:
                    return float(x)
                except:
                    return 0.0

            ch = f(cen_h)
            te = f(top_e)
            be = f(bot_e)
            ot = int(off_t or 0)
            ob = int(off_b or 0)
            try:
                eid = d.element.Id.IntegerValue
            except:
                eid = d.element.Id

            output.print_md(
                "{} | {} | CLh:{:.0f}\" | T:{}\" B:{}\" | ΔT:{:+.0f}\" ΔB:{:+.0f}\" | {}".format(
                    eid, family, ch, ot, ob, te, be, tag
                )
            )
    t.Commit()

    dbg("PARAM CHECK: width_in={}, heigth_in={}, width_out={}, heigth_out={}",
        d.width_in, d.heigth_in, d.width_out, d.heigth_out)
