# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

# Standard library
# =========================================================
from revit_xyz import RevitXYZ
from Autodesk.Revit.DB import (
    ElementId,
    FilteredElementCollector,
    BuiltInCategory,
    UnitUtils,
    FabricationPart,
    UnitTypeId,
    ConnectorType
)
import re
import logging
import math
from enum import Enum

# Thrid Party
from pyrevit import DB, revit, script

#
import clr
clr.AddReference("RevitAPI")

# Variables
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
view = revit.active_view
output = script.get_output()

# Logging
log = logging.getLogger("RevitDuct")


# Revut Duct Class
# ============================================================
class RevitOffset:
    def __init__(self, doc, view, element):
        self.doc = doc
        self.view = view
        self.element = element

    @property
    def offset_data(self):
        """Cache and return offset calculations for the duct."""
        if not hasattr(self, '_offset_data'):
            # Use identified inlet/outlet instead of raw connectors
            c_in, c_out = self.identify_inlet_outlet()

            if c_in and c_out:
                # Detect round connectors (prefer explicit connector properties)
                def has_radius(conn):
                    try:
                        return hasattr(conn, 'Radius') and conn.Radius and conn.Radius > 1e-6
                    except Exception:
                        return False

                is_round_in = has_radius(c_in)
                is_round_out = has_radius(c_out)
                is_round = bool(is_round_in and is_round_out)

                # Get dimensions based on shape
                if is_round:
                    # For round: use diameter from connector or parameters
                    w_i = None
                    w_o = None
                    try:
                        r_in = c_in.Radius
                        if r_in and r_in > 1e-6:
                            w_i = r_in * 24.0
                    except Exception:
                        pass
                    if not w_i:
                        w_i = self.diameter_in

                    try:
                        r_out = c_out.Radius
                        if r_out and r_out > 1e-6:
                            w_o = r_out * 24.0
                    except Exception:
                        pass
                    if not w_o:
                        w_o = self.diameter_out
                    if not w_o:
                        w_o = w_i

                    h_i = w_i
                    h_o = w_o
                else:
                    # For rectangular: use width/height parameters
                    w_i = self.width_in
                    h_i = self.heigth_in
                    w_o = self.width_out or w_i
                    h_o = self.heigth_out or h_i

                # Validate we have dimensions
                if not w_i or not h_i:
                    self._offset_data = None
                    return self._offset_data

                # Revit internal units (feet) -> inches
                p_in = (c_in.Origin.X * 12.0, c_in.Origin.Y *
                        12.0, c_in.Origin.Z * 12.0)
                p_out = (c_out.Origin.X * 12.0, c_out.Origin.Y *
                         12.0, c_out.Origin.Z * 12.0)

                # Get coordinate system from INLET (cache to avoid repeated access)
                try:
                    cs = c_in.CoordinateSystem
                    bx = cs.BasisX
                    by = cs.BasisY
                    u_hat = (bx.X, bx.Y, bx.Z)
                    v_hat = (by.X, by.Y, by.Z)
                except Exception:
                    u_hat = (1.0, 0.0, 0.0)
                    v_hat = (0.0, 1.0, 0.0)

                # Keep height axis pointing up in world space to stabilize
                # top/bottom
                if v_hat[2] < 0.0:
                    u_hat = (-u_hat[0], -u_hat[1], -u_hat[2])
                    v_hat = (-v_hat[0], -v_hat[1], -v_hat[2])

                # Centerline offsets (inlet to outlet)
                delta = (
                    p_out[0] - p_in[0], p_out[1] - p_in[1], p_out[2] - p_in[2])
                width_offset = abs(RevitXYZ.dot(delta, u_hat))
                height_offset = abs(RevitXYZ.dot(delta, v_hat))

                # Edge offsets (inlet to outlet)
                if not is_round:
                    edge_offsets = RevitXYZ.edge_diffs_whole_in(
                        p_in, w_i, h_i, p_out, w_o, h_o, u_hat, v_hat)
                else:
                    # Round parts do not have meaningful rectangular edges.
                    # Provide None for edge offsets and include diameters for context.
                    try:
                        d_in = c_in.Radius * 24.0
                    except Exception:
                        d_in = self.diameter_in
                    try:
                        d_out = c_out.Radius * 24.0
                    except Exception:
                        d_out = self.diameter_out or d_in
                    edge_offsets = {
                        'whole_in': {
                            'left': None,
                            'right': None,
                            'top': None,
                            'bottom': None,
                        },
                        'round': True,
                        'diam_in': d_in,
                        'diam_out': d_out,
                    }

                self._offset_data = {
                    'centerline_width': width_offset,
                    'centerline_height': height_offset,
                    'edges': edge_offsets
                }
            else:
                # No valid connectors: cannot compute geometry-based offsets
                self._offset_data = None

        return self._offset_data

    @property
    def is_round(self):
        """True if both connectors are round (edge offsets not meaningful)."""
        data = getattr(self, '_offset_data', None)
        if not data:
            # Force calculation once if missing
            data = self.offset_data
        return bool(data and data.get('edges') and data['edges'].get('round'))

    @property
    def centerline_width(self):
        """Centerline width offset in inches."""
        data = self.offset_data
        return data['centerline_width'] if data else None

    @property
    def centerline_height(self):
        """Centerline height offset in inches."""
        data = self.offset_data
        return data['centerline_height'] if data else None

    @property
    def offset_left(self):
        """Left edge offset in whole inches."""
        data = self.offset_data
        return data['edges']['whole_in']['left'] if data and data['edges'] else None

    @property
    def offset_right(self):
        """Right edge offset in whole inches."""
        data = self.offset_data
        return data['edges']['whole_in']['right'] if data and data['edges'] else None

    @property
    def offset_top(self):
        """Top edge offset in whole inches."""
        data = self.offset_data
        return data['edges']['whole_in']['top'] if data and data['edges'] else None

    @property
    def offset_bottom(self):
        """Bottom edge offset in whole inches."""
        data = self.offset_data
        if data and data['edges']:
            return data['edges']['whole_in']['bottom']
        return None

    def identify_inlet_outlet(self):
        """Deterministically pick inlet (larger connector) and outlet (smaller)."""
        try:
            conns = list(self.element.ConnectorManager.Connectors)
            if len(conns) < 2:
                return (None, None)
            c0, c1 = conns[0], conns[1]

            # Try rectangular sizes (inches)
            def rect_wh(conn):
                try:
                    return conn.Width * 12.0, conn.Height * 12.0
                except Exception:
                    return None, None

            w0, h0 = rect_wh(c0)
            w1, h1 = rect_wh(c1)

            # Try round diameters (inches)
            def diameter(conn):
                try:
                    return conn.Radius * 24.0  # 2 * radius * 12
                except Exception:
                    return None
            d0 = diameter(c0)
            d1 = diameter(c1)

            # Rectangular case first
            if w0 and h0 and w1 and h1:
                a0 = w0 * h0
                a1 = w1 * h1
                if abs(a0 - a1) > 1e-6:
                    return (c0, c1) if a0 >= a1 else (c1, c0)
                # Tie: fall back to element id for stability
                id0 = get_element_id_value(c0.Owner.Id)
                id1 = get_element_id_value(c1.Owner.Id)
                return (c0, c1) if id0 <= id1 else (c1, c0)

            # Round case
            if d0 and d1:
                if abs(d0 - d1) > 1e-6:
                    return (c0, c1) if d0 >= d1 else (c1, c0)
                id0 = get_element_id_value(c0.Owner.Id)
                id1 = get_element_id_value(c1.Owner.Id)
                return (c0, c1) if id0 <= id1 else (c1, c0)

            # Mixed shape case: one rectangular, one round
            # Compute areas and compare
            a0 = None
            a1 = None
            if w0 and h0:
                a0 = w0 * h0
            elif d0:
                a0 = 3.14159 * (d0 / 2.0) ** 2

            if w1 and h1:
                a1 = w1 * h1
            elif d1:
                a1 = 3.14159 * (d1 / 2.0) ** 2

            if a0 is not None and a1 is not None and abs(a0 - a1) > 1e-6:
                return (c0, c1) if a0 >= a1 else (c1, c0)

            # Last resort: use connector index for deterministic ordering
            # (both connectors belong to same element, so Owner.Id won't help)
            return (c0, c1)  # c0 is always inlet for consistency

        except Exception:
            return (None, None)

    def classify_offset(self):
        """Classify transition/reducer offset as CL/FOB/FOT/FOS or arrow/numeric offset."""
        c_in, c_out = self.identify_inlet_outlet()
        if not (c_in and c_out):
            return None

        p_in = c_in.Origin
        p_out = c_out.Origin

        # Vector from inlet to outlet
        delta = p_out - p_in

        # Get width direction (BasisX) from inlet connector
        # BasisX points along the width of rectangular duct (left-right direction)
        try:
            width_dir = c_in.CoordinateSystem.BasisX
            # Project delta onto width direction to get signed horizontal offset
            # Positive = offset in +BasisX direction (right), Negative = offset in -BasisX direction (left)
            offset_perp_signed = (delta.X * width_dir.X +
                                  delta.Y * width_dir.Y +
                                  delta.Z * width_dir.Z) * 12.0
            offset_perp = abs(offset_perp_signed)
        except Exception:
            # Fallback: no horizontal offset
            offset_perp_signed = 0.0
            offset_perp = 0.0
            offset_perp = 0.0

        # Horizontal centerline offset (plan distance - for reference)
        cen_w = math.hypot(delta.X, delta.Y) * 12

        # Vertical centerline offset (signed and magnitude)
        cen_h_signed = delta.Z * 12.0
        cen_h = abs(cen_h_signed)

        # Detect connector shapes
        def is_round(conn):
            try:
                return hasattr(conn, 'Radius') and conn.Radius and conn.Radius > 1e-6
            except Exception:
                return False

        round_in = is_round(c_in)
        round_out = is_round(c_out)

        # Handle mixed transitions (square to round)
        if round_in != round_out:
            # Mixed transition: use centerline offsets and perpendicular offset
            return {
                'centerline_w': cen_w,
                'centerline_h': cen_h,
                'centerline_h_signed': delta.Z * 12.0,
                'offset_perp': offset_perp,
                'offset_perp_signed': offset_perp_signed,
                'top_edge': None,
                'bot_edge': None,
                'left_edge': None,
                'right_edge': None,
                'top_mag': None,
                'bot_mag': None,
                'left_mag': None,
                'right_mag': None,
                'top_aligned': False,
                'bot_aligned': False,
                'left_aligned': False,
                'right_aligned': False,
                'cl_vert': True,
                'is_mixed': True
            }

        # Sizes (both width and height)
        w_in = c_in.Width * \
            12.0 if hasattr(c_in, 'Width') and c_in.Width else 0.0
        w_out = c_out.Width * \
            12.0 if hasattr(c_out, 'Width') and c_out.Width else w_in

        h_in = (c_in.Height * 12.0 if hasattr(c_in, 'Height') and c_in.Height
                else (c_in.Radius * 24.0 if hasattr(c_in, 'Radius') and c_in.Radius else 0.0))
        h_out = (c_out.Height * 12.0 if hasattr(c_out, 'Height') and c_out.Height
                 else (c_out.Radius * 24.0 if hasattr(c_out, 'Radius') and c_out.Radius else h_in))

        # World Z planes (feet) - using actual connector positions
        top_in_z = p_in.Z + 0.5 * (h_in / 12.0)
        top_out_z = p_out.Z + 0.5 * (h_out / 12.0)
        bot_in_z = p_in.Z - 0.5 * (h_in / 12.0)
        bot_out_z = p_out.Z - 0.5 * (h_out / 12.0)

        # Edge rises (inches, signed)
        top_e = (top_out_z - top_in_z) * 12.0
        bot_e = (bot_out_z - bot_in_z) * 12.0

        # Left and right edge offsets (if rectangular)
        left_in_z = p_in.Z - 0.5 * (w_in / 12.0)
        left_out_z = p_out.Z - 0.5 * (w_out / 12.0)
        right_in_z = p_in.Z + 0.5 * (w_in / 12.0)
        right_out_z = p_out.Z + 0.5 * (w_out / 12.0)

        left_e = (left_out_z - left_in_z) * 12.0
        right_e = (right_out_z - right_in_z) * 12.0

        # Tolerance
        tol_in = 0.01

        top_aligned = abs(top_e) < tol_in
        bot_aligned = abs(bot_e) < tol_in
        left_aligned = abs(left_e) < tol_in
        right_aligned = abs(right_e) < tol_in
        cl_vert = top_aligned and bot_aligned

        # Whole-inch magnitudes
        off_t = int(round(abs(top_e)))
        off_b = int(round(abs(bot_e)))
        off_l = int(round(abs(left_e)))
        off_r = int(round(abs(right_e)))

        return {
            'centerline_w': cen_w,
            'centerline_h': cen_h,
            'centerline_h_signed': cen_h_signed,
            'offset_perp': offset_perp,
            'offset_perp_signed': offset_perp_signed,
            'top_edge': top_e,
            'bot_edge': bot_e,
            'left_edge': left_e,
            'right_edge': right_e,
            'top_mag': off_t,
            'bot_mag': off_b,
            'left_mag': off_l,
            'right_mag': off_r,
            'top_aligned': top_aligned,
            'bot_aligned': bot_aligned,
            'left_aligned': left_aligned,
            'right_aligned': right_aligned,
            'cl_vert': cl_vert,
            'is_mixed': False,
            'w_in': w_in,
            'w_out': w_out,
            'h_in': h_in,
            'h_out': h_out
        }

    def get_offset_value(self):
        """Calculate offset classification tag for transitions/reducers/offsets.

        Returns:
            str: Tag like "CL", "FOB", "FOT", "FOS", "↑2"", "3"→", or None if not applicable.
        """
        family = (self.family or "").lower().strip()

        # Family lists
        reducer_square = ["transition"]
        reducer_round = ["reducer"]
        offset_list = ["ogee", "offset", "radius offset",
                       "mitered offset", "mitred offset"]
        family_list = reducer_square + reducer_round + offset_list

        if family not in family_list:
            return None

        # Get offset data
        offset_data = self.classify_offset()
        if not offset_data:
            return None

        cen_w = offset_data['centerline_w']
        cen_h = offset_data['centerline_h']
        top_e = offset_data['top_edge']
        bot_e = offset_data['bot_edge']
        top_aligned = offset_data['top_aligned']
        bot_aligned = offset_data['bot_aligned']
        cl_vert = offset_data['cl_vert']

        # Rectangular reducers/transitions
        if family in reducer_square:
            is_rotation = (cen_h < 0.5) and abs(abs(top_e) - abs(bot_e)) < 0.5

            # Get left/right edge data
            left_e = offset_data.get('left_edge', 0)
            right_e = offset_data.get('right_edge', 0)
            left_aligned = offset_data.get('left_aligned', False)
            right_aligned = offset_data.get('right_aligned', False)

            if cl_vert or is_rotation:
                return "CL"

            # Build combined tag for aligned edges
            aligned_edges = []
            if bot_aligned:
                aligned_edges.append("FOB")
            if top_aligned:
                aligned_edges.append("FOT")
            if left_aligned:
                aligned_edges.append("FOL")
            if right_aligned:
                aligned_edges.append("FOR")

            if aligned_edges:
                return "/".join(aligned_edges)

            # No edges aligned - show offsets with arrows
            # Check if both vertical AND horizontal offsets exist
            has_vert = abs(top_e) >= 0.5
            has_horiz = abs(left_e) >= 0.5 or abs(right_e) >= 0.5

            if has_vert and has_horiz:
                # Both directions - show both with space
                vert_mag = int(round(abs(top_e)))
                horiz_mag = int(round(abs(left_e)))
                vert_str = u'↑{}"TU'.format(
                    vert_mag) if top_e > 0 else u'↓{}"TD'.format(vert_mag)
                horiz_str = u'←{}"'.format(
                    horiz_mag) if left_e < 0 else u'→{}"'.format(horiz_mag)
                return u'{} {}'.format(vert_str, horiz_str)
            elif has_vert:
                # Only vertical
                mag = int(round(abs(top_e)))
                return u'↑{}"TU'.format(mag) if top_e > 0 else u'↓{}"TD'.format(mag)
            elif has_horiz:
                # Only horizontal
                mag = int(round(abs(left_e)))
                return u'←{}"'.format(mag) if left_e < 0 else u'→{}"'.format(mag)
            else:
                return "CL"

        # Round reducers
        elif family in reducer_round:
            y_off = self.reducer_offset
            d_in = self.diameter_in
            d_out = self.diameter_out

            if (y_off is not None) and (d_in is not None) and (d_out is not None):
                expected_cl = (d_in - d_out) / 2.0

                if abs(y_off - expected_cl) < 0.01:
                    return "CL"
                elif abs(d_out + y_off - d_in) < 0.01 or abs(y_off) < 0.1:
                    return "FOS"
                else:
                    return u'{}"→'.format(abs(int(round(y_off))))

        # Horizontal offsets
        elif family in offset_list:
            oge_o = self.ogee_offset
            offset = oge_o or cen_w or 0
            return u'{}"→'.format(int(round(offset)))

        return None

    def get_connected_elements(self, connector_index=0):
        """Gets all elements connected to the selected element"""
        connector = self.get_connector(connector_index)
        connected_elements = []

        if connector and connector.IsConnected:
            for ref_conn in connector.AllRefs:
                if ref_conn.Owner.Id != self.element.Id:
                    connected_elements.append(ref_conn.Owner)
        return connected_elements

    def trace_run(start_duct, seen_ids, allowed_duct):
        """Recursively follow connections to build a complete run"""
        run = []  # Store ducts in this run
        stack = [start_duct]  # Ducts to process

        while stack:
            current = stack.pop()
            if current.Id.IntegerValue in seen_ids:
                continue

            run.append(current)
            seen_ids.add(current.Id.IntegerValue)

            # Follow all connections
            for connector_index in [0, 1, 2]:
                connected = current.get_connected_elements(connector_index)
                for elem in connected:
                    duct = RevitDuct(doc, view, elem)
                    if duct.family and duct.family.strip().lower() in allowed_duct:
                        stack.append(elem)

        return run
