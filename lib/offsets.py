# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

# Imports
# =========================================================================
from size import Size
from revit_xyz import RevitXYZ
from Points import Points
from xyz import XYZ as CustomXYZ

# Class
# =========================================================================


class Offsets:
    def __init__(self, element):
        self.element = element
        self.xyz = RevitXYZ(element)
        self.error_msg = None

        # Get start and end points preferring connector origins (like spam script)
        self.start_point, self.end_point = self._get_inlet_outlet_points()

        # Get size from element parameter
        size_param = element.LookupParameter("Size")
        size_string = size_param.AsString() if size_param else None
        self.size = Size(size_string) if size_string else None

        # Build error message if data is missing
        if not self.start_point or not self.end_point:
            self.error_msg = "Element has no Location.Curve (not a linear element)"
        elif not size_string:
            self.error_msg = "Element has no 'Size' parameter"
        elif self.size and (not self.size.in_size or not self.size.out_size):
            self.error_msg = "Size parameter could not be parsed: '{}'".format(
                size_string)

    def calculate_offsets(self):
        if not self.start_point or not self.end_point or not self.size:
            return None

        # Convert Revit XYZ to custom XYZ for Points class
        inlet = CustomXYZ(self.start_point.X * 12,
                          self.start_point.Y * 12,
                          self.start_point.Z * 12)

        outlet = CustomXYZ(self.end_point.X * 12,
                           self.end_point.Y * 12,
                           self.end_point.Z * 12)

        # Calculate direction and perpendicular vectors
        direction = (outlet - inlet).normalize()

        if abs(direction.Z) < 0.9:
            global_up = CustomXYZ(0, 0, 1)
        else:
            global_up = CustomXYZ(1, 0, 0)

        right = direction.cross(global_up).normalize()
        up = right.cross(direction).normalize()

        # Create Points objects with right and up vectors
        inlet_points = Points(inlet, outlet, right, up)
        outlet_points = Points(inlet, outlet, right, up)

        # Generate perimeter points based on shape
        in_shape = self.size.in_shape()
        if in_shape == "round":
            in_perimeter = inlet_points.round(self.size.in_diameter)
        elif in_shape == "oval":
            in_perimeter = inlet_points.oval(
                self.size.in_width, self.size.in_height)
        elif in_shape == "rectangle":
            in_perimeter = inlet_points.rectangle(
                self.size.in_width, self.size.in_height)
        else:
            return None

        out_shape = self.size.out_shape()
        if out_shape == "round":
            out_perimeter = outlet_points.round(
                self.size.out_diameter
            )

        elif out_shape == "oval":
            out_perimeter = outlet_points.oval(
                self.size.out_width,
                self.size.out_height
            )

        elif out_shape == "rectangle":
            out_perimeter = outlet_points.rectangle(
                self.size.out_width,
                self.size.out_height
            )

        else:
            return None

        top_in = in_perimeter[0]  # 0°
        right_in = in_perimeter[90]  # 90°
        bottom_in = in_perimeter[180]  # 180°
        left_in = in_perimeter[270]  # 270°

        top_out = out_perimeter[0]
        right_out = out_perimeter[90]
        bottom_out = out_perimeter[180]
        left_out = out_perimeter[270]

        # Calculate offset vectors
        top_offset_vec = top_out - top_in
        right_offset_vec = right_out - right_in
        bottom_offset_vec = bottom_out - bottom_in
        left_offset_vec = left_out - left_in

        # Project onto the right and up directions to get component offsets
        top_val = top_offset_vec.dot(up)
        bottom_val = bottom_offset_vec.dot(up)
        right_val = right_offset_vec.dot(right)
        left_val = left_offset_vec.dot(right)

        offsets = {
            'top': top_val,
            'right': right_val,
            'bottom': bottom_val,
            'left': left_val,
            # Center shift is the mean of opposite edges, not a duplicate of one edge
            'center_horizontal': (right_val + left_val) / 2.0,
            'center_vertical': (top_val + bottom_val) / 2.0,
        }

        return offsets

    def _get_inlet_outlet_points(self):
        """Return (inlet, outlet) XYZs using connector origins when available,
        mirroring the spam script approach. Fallback to Location.Curve endpoints.
        """
        # Try connectors first
        pts = []
        seen = set()

        def add_xyz(o):
            if o:
                key = (round(o.X, 9), round(o.Y, 9), round(o.Z, 9))
                if key not in seen:
                    seen.add(key)
                    pts.append(o)

        try:
            cm = getattr(self.element, 'ConnectorManager', None)
            connectors = cm.Connectors if cm else getattr(self.element, 'Connectors', None)
            if connectors:
                count = getattr(connectors, 'Size', getattr(connectors, 'Count', 0))
                if count and hasattr(connectors, 'Item'):
                    for i in range(count):
                        add_xyz(getattr(connectors.Item(i), 'Origin', None))
                try:
                    for c in connectors:
                        add_xyz(getattr(c, 'Origin', None))
                except Exception:
                    pass

            # Primary/Secondary connectors
            pc = getattr(self.element, 'PrimaryConnector', None)
            sc = getattr(self.element, 'SecondaryConnector', None)
            add_xyz(getattr(pc, 'Origin', None))
            add_xyz(getattr(sc, 'Origin', None))

            # GetConnectors API (fabrication)
            get_conns = getattr(self.element, 'GetConnectors', None)
            if get_conns:
                try:
                    for c in get_conns() or []:
                        add_xyz(getattr(c, 'Origin', None))
                except Exception:
                    pass
        except Exception:
            pass

        # If more than two points, choose the pair with max distance to define inlet/outlet
        if len(pts) > 2:
            max_d2 = -1.0
            a, b = 0, 1
            for i in range(len(pts)):
                for j in range(i + 1, len(pts)):
                    d2 = (pts[i].X - pts[j].X) ** 2 + (pts[i].Y - pts[j].Y) ** 2 + (pts[i].Z - pts[j].Z) ** 2
                    if d2 > max_d2:
                        max_d2 = d2
                        a, b = i, j
            pts = [pts[a], pts[b]]

        # If we found two distinct connector points, use them directly
        if len(pts) >= 2:
            # Return in the order given; Offsets is symmetric and uses direction defined by these
            return pts[0], pts[1]

        # If only one connector point, pair it with curve endpoints to create distinct points
        loc = getattr(self.element, 'Location', None)
        curve = getattr(loc, 'Curve', None) if loc else None
        if len(pts) == 1 and curve:
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            c = pts[0]
            # Closest to start becomes inlet; other becomes outlet
            inlet = c if (c - p0).GetLength() <= (c - p1).GetLength() else p0
            outlet = p1 if inlet == c else c
            return inlet, outlet

        # Fallback: use curve endpoints if available
        if curve:
            return curve.GetEndPoint(0), curve.GetEndPoint(1)

        # Nothing found
        return None, None


if __name__ == "__main__":
    # Test data
    start = (8.115217245, -20.102905699, 10.500000000)
    end = (8.115217245, -12.181281631, 10.500000000)
    shape = "40/20-12Ø"

    print("Test case: {}".format(shape))
    print("Start point (inches): {}".format(start))
    print("End point (inches): {}".format(end))
    print("\nNote: This is a test stub. To use with actual Revit elements:")
    print("  element = [your Revit element]")
    print("  offsets = Offsets(element)")
    print("  result = offsets.calculate_offsets()")
