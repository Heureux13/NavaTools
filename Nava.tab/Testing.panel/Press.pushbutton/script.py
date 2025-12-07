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
__title__ = "Press"
__doc__ = """
Assigns offset information about specific duct fittings
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

family_list = {
    "transition", "mitred offset",
    "radius offset", "mitered offset",
    "ogee", "offset", "reducer", "square to ø"
}

reducer_square = {
    "transition"
}

reducer_round = {
    "reducer"
}

square_round = {
    "square to ø"
}

offset_list = {
    "ogee", "offset", "radius offset",
    "mitered offset", "mitred offset"
}


with Transaction(doc, "Offset Parameter") as t:
    t.Start()

    for d in ducts:
        tag = None
        family = d.family.lower().strip()
        output.print_md(
            "Family: {} | Size: {} | Length: {:06.2f}".format(
                d.family,
                d.size,
                d.length
            )
        )

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
            # Get offset geometry data
            offset_info = d.offset_data

            # Optional debug
            dbg("cen_w: {:.2f} in, cen_h: {:.2f} in", cen_w, cen_h)
            dbg("RAW centerline_h_signed from offset_data: {:.2f}",
                offset_data.get('centerline_h_signed', 'N/A'))
            dbg("top_e: {:+.2f}\" bot_e: {:+.2f}\"", top_e, bot_e)
            top_al = bool(offset_data.get('top_aligned', False))
            bot_al = bool(offset_data.get('bot_aligned', False))
            right_al = bool(offset_data.get('right_aligned', False))
            left_al = bool(offset_data.get('left_aligned', False))
            v_is_cl = top_al and bot_al
            h_is_cl = right_al and left_al
            # Horizontal signed plan-perpendicular offset for FOR/FOL mapping
            h_signed = float(offset_data.get('offset_perp_signed', 0) or 0)
            EPS = 0.49
            dbg("top_al:{} bot_al:{} right_al:{} left_al:{} | vCL:{} hCL:{}",
                top_al,
                bot_al,
                right_al,
                left_al,
                v_is_cl,
                h_is_cl
                )
            # Tag priority: CL -> single-side flush -> combos -> else numeric
            # Exception: offset families skip CL/flush checks (alignment flags not meaningful)
            tag = None
            is_rotation = False
            if top_e is not None and bot_e is not None:
                is_rotation = (cen_h < 0.5) and abs(
                    abs(top_e) - abs(bot_e)) < 0.5

            # For mixed shapes (square-to-round), check actual magnitudes not cl_vert flag
            is_actually_cl = False
            if family in square_round:
                v_signed_check = float(offset_data.get(
                    'centerline_h_signed', 0) or 0)
                is_actually_cl = (
                    abs(v_signed_check) < EPS and abs(h_signed) < EPS)

            if family not in offset_list:
                if family in square_round:
                    if is_actually_cl:
                        tag = "CL"
                elif (v_is_cl and h_is_cl) or (cl_vert or is_rotation):
                    tag = "CL"

            # If exactly one side is CL, emit the non-CL side label only (skip for offsets)
            if tag is None and family not in offset_list and family not in square_round:
                if v_is_cl and not h_is_cl:
                    # Derive horizontal flush label from plan-perp sign
                    if h_signed > EPS:
                        tag = 'FOR'
                    elif h_signed < -EPS:
                        tag = 'FOL'
                elif h_is_cl and not v_is_cl:
                    if top_al and not bot_al:
                        tag = 'FOT'
                    elif bot_al and not top_al:
                        tag = 'FOB'

            # If any side is flush to its edge (not CL), allow combos (skip for offsets)
            if tag is None and family not in offset_list:
                labels = []
                if family in square_round:
                    # Square-to-round: label which edge of the inlet has clearance
                    # Only apply flush labels when there's BOTH vertical and horizontal offset
                    # Pure horizontal offsets should use numeric format instead
                    v_signed_check = float(offset_data.get(
                        'centerline_h_signed', 0) or 0)
                    dbg("Square-to-round | v_signed:{:.2f} h_signed:{:.2f}",
                        v_signed_check, h_signed)

                    has_vertical = abs(v_signed_check) > EPS
                    has_horizontal = abs(h_signed) > EPS

                    # Only use flush labels if there's both vertical AND horizontal offset
                    # Pure H or V offsets fall through to numeric composer
                    if has_vertical and has_horizontal:
                        if v_signed_check < -EPS:
                            # Outlet down = bottom clearance
                            labels.append('FOB')
                        elif v_signed_check > EPS:
                            labels.append('FOT')  # Outlet up = top clearance
                        if h_signed > EPS:
                            labels.append('FOR')  # Positive BasisX = FOR
                        elif h_signed < -EPS:
                            labels.append('FOL')  # Negative BasisX = FOL
                    dbg("Square-to-round labels: {}", labels)
                else:
                    # Use alignment flags for rect/round shapes
                    if not v_is_cl:
                        if top_al:
                            labels.append('FOT')
                        if bot_al:
                            labels.append('FOB')
                    if not h_is_cl:
                        # Use h_signed to decide FOR/FOL
                        if h_signed > EPS:
                            labels.append('FOR')
                        elif h_signed < -EPS:
                            labels.append('FOL')
                if labels:
                    tag = ' | '.join(labels)

            # 3) FOS (round reducers specific)
            if tag is None and family in reducer_round:
                y_off = d.reducer_offset
                d_in = d.diameter_in
                d_out = d.diameter_out
                if y_off is not None and d_in and d_out:
                    expected_cl = (d_in - d_out) / 2.0
                    if abs(y_off - expected_cl) < 0.01:
                        tag = "CL"
                    elif abs(d_out + y_off - d_in) < 0.01 or abs(y_off) < 0.1:
                        tag = "FOS"

            # 4) Fall back to combined V|H (no special label)
            # Leave tag as None to trigger combined composer below

            # Compose combined vertical | horizontal components when no special tag was set
            if tag is None:
                # For vertical: use top edge for rect/round, centerline_h for square_round
                if family in square_round:
                    v_signed = float(offset_data.get(
                        'centerline_h_signed', 0) or 0)
                else:
                    v_signed = float(offset_data.get('top_edge', 0) or 0)
                h_signed = float(offset_data.get('offset_perp_signed', 0) or 0)

                # Use a tolerance aligned with inch rounding to avoid -0"
                EPS = 0.49
                v_abs = abs(v_signed)
                h_abs = abs(h_signed)
                vmag = int(round(v_abs))
                hmag = int(round(h_abs))

                # If both components are effectively zero, mark centerline
                if vmag == 0 and hmag == 0:
                    tag = "CL"
                    dbg(
                        "Combined tag | both ~0 => CL (V:{:.2f}, H:{:.2f})", v_signed, h_signed)
                elif vmag == 0 and hmag != 0:
                    # Only horizontal component present
                    hdir = 'L' if h_signed > EPS else (
                        'R' if h_signed < -EPS else '')
                    tag = '{}"{}'.format(
                        hmag, hdir) if hdir else '{}"'.format(hmag)
                    dbg("Combined tag | only H => {} (H:{:.2f})", tag, h_signed)
                elif hmag == 0 and vmag != 0:
                    # Only vertical component present
                    vdir = 'UP' if v_signed > EPS else (
                        'DW' if v_signed < -EPS else '')
                    tag = '{}"{}'.format(
                        vmag, vdir) if vdir else '{}"'.format(vmag)
                    dbg("Combined tag | only V => {} (V:{:.2f})", tag, v_signed)
                else:
                    vdir = 'UP' if v_signed > EPS else (
                        'DW' if v_signed < -EPS else '')
                    hdir = 'L' if h_signed > EPS else (
                        'R' if h_signed < -EPS else '')

                    vert_part = '{}"{}'.format(
                        vmag, vdir) if vdir else '{}"'.format(vmag)
                    horiz_part = '{}"{}'.format(
                        hmag, hdir) if hdir else '{}"'.format(hmag)
                    tag = '{} | {}'.format(vert_part, horiz_part)
                    dbg("Combined tag | V:{:.2f} ({} -> {}) H:{:.2f} ({} -> {}) => {}",
                        v_signed, vdir or '0', vmag, h_signed, hdir or '0', hmag, tag)
            # else: tag was already set (CL/FOT/FOB/FOR/FOL/FOS combos)

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
