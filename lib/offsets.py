# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

# Imports
# =========================================================================
from points import Points
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

        # Use connector basis vectors if available, otherwise calculate from direction
        if inlet_basis_x and inlet_basis_y:
            right = XYZ(inlet_basis_x.X, inlet_basis_x.Y, inlet_basis_x.Z)
            up = XYZ(inlet_basis_y.X, inlet_basis_y.Y, inlet_basis_y.Z)
        else:
            # Fallback: calculate from direction
            direction = (outlet_pt - inlet_pt).normalize()
            if abs(direction.Z) < 0.9:
                global_up = XYZ(0, 0, 1)
            else:
                global_up = XYZ(1, 0, 0)
            right = direction.cross(global_up).normalize()
            up = right.cross(direction).normalize()

        # Create Points objects with right and up vectors
        inlet_points = Points(inlet_pt, outlet_pt, right, up)
        outlet_points = Points(outlet_pt, inlet_pt, right, up)

        # Generate inlet perimeter based on shape
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

        # Generate outlet perimeter based on shape
        out_shape = self.size.out_shape()
        if out_shape == "round":
            out_perimeter = outlet_points.round(self.size.out_diameter)
        elif out_shape == "oval":
            out_perimeter = outlet_points.oval(
                self.size.out_width, self.size.out_height)
        elif out_shape == "rectangle":
            out_perimeter = outlet_points.rectangle(
                self.size.out_width, self.size.out_height)
        else:
            return None

        # Get cardinal points
        top_in = in_perimeter[0]
        right_in = in_perimeter[90]
        bottom_in = in_perimeter[180]
        left_in = in_perimeter[270]

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
            'center_horizontal': (right_val + left_val) / 2.0,
            'center_vertical': (top_val + bottom_val) / 2.0,
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
