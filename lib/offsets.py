# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

# Imports
# =========================================================================
from xyz import XYZ

# Class
# =========================================================================


class Offsets:
    def __init__(self, inlet_data, outlet_data, size):
        """Pure offset calculation class.

        Args:
            inlet_data: Dict with 'origin', 'basis_x', 'basis_y', 'basis_z'
            outlet_data: Dict with 'origin', 'basis_x', 'basis_y', 'basis_z'
            size: Size object with parsed inlet/outlet dimensions
        """
        self.inlet_data = inlet_data
        self.outlet_data = outlet_data
        self.size = size

    def calculate(self):
        """Calculate offset vectors between inlet and outlet perimeters."""
        if not self.inlet_data or not self.outlet_data or not self.size:
            return None

        inlet_origin = self.inlet_data.get('origin')
        outlet_origin = self.outlet_data.get('origin')
        inlet_basis_x = self.inlet_data.get('basis_x')
        inlet_basis_y = self.inlet_data.get('basis_y')

        if not inlet_origin or not outlet_origin:
            return None

        # Convert Revit XYZ to custom XYZ for Points class (feet to inches)
        inlet_pt = XYZ(
            inlet_origin.X * 12, inlet_origin.Y * 12, inlet_origin.Z * 12)
        outlet_pt = XYZ(
            outlet_origin.X * 12, outlet_origin.Y * 12, outlet_origin.Z * 12)

        # Inlet orientation (basis if present, else derived)
        if inlet_basis_x and inlet_basis_y:
            right_in = XYZ(
                inlet_basis_x.X,
                inlet_basis_x.Y,
                inlet_basis_x.Z
            ).normalize()
            up_in = XYZ(
                inlet_basis_y.X,
                inlet_basis_y.Y,
                inlet_basis_y.Z
            ).normalize()
        else:
            direction = (outlet_pt - inlet_pt).normalize()
            global_up = XYZ(0, 0, 1) if abs(
                direction.Z) < 0.9 else XYZ(1, 0, 0)
            right_in = direction.cross(global_up).normalize()
            up_in = right_in.cross(direction).normalize()

        # Half-sizes (inches)
        def halves(shape, w, h, d):
            if shape == "round":
                r = (d or 0) / 2.0
                return r, r
            if shape == "oval":
                return (w or 0) / 2.0, (h or 0) / 2.0
            return (w or 0) / 2.0, (h or 0) / 2.0

        in_shape = self.size.in_shape()
        out_shape = self.size.out_shape()
        in_half_w, in_half_h = halves(
            in_shape, self.size.in_width, self.size.in_height, self.size.in_diameter)
        out_half_w, out_half_h = halves(
            out_shape, self.size.out_width, self.size.out_height, self.size.out_diameter)

        # Center displacement
        center_vec = outlet_pt - inlet_pt

        # Offsets: center movement projected on right/up, plus size deltas
        center_h = center_vec.dot(right_in)
        center_v = center_vec.dot(up_in)

        top_val = center_v + (out_half_h - in_half_h)
        bottom_val = center_v - (out_half_h - in_half_h)
        right_val = center_h + (out_half_w - in_half_w)
        left_val = center_h - (out_half_w - in_half_w)

        offsets = {
            'top': top_val,
            'right': right_val,
            'bottom': bottom_val,
            'left': left_val,
            'center_horizontal': center_h,
            'center_vertical': center_v,
        }

        return offsets


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
