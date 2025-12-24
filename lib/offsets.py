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

        # Determine if duct is rotated (width/height may be swapped)
        # Check which dimension of the duct aligns with vertical
        if inlet_basis_x and inlet_basis_y:
            bx_z = abs(inlet_basis_x.Z)
            by_z = abs(inlet_basis_y.Z)

            # If basis_x is more vertical, the duct is rotated 90° (width is vertical)
            if bx_z > by_z and bx_z > 0.5:
                # Duct rotated - swap width/height
                in_w = self.size.in_height if self.size.in_height else 0
                in_h = self.size.in_width if self.size.in_width else 0
                out_w = self.size.out_height if self.size.out_height else 0
                out_h = self.size.out_width if self.size.out_width else 0
            else:
                # Normal orientation
                in_w = self.size.in_width if self.size.in_width else 0
                in_h = self.size.in_height if self.size.in_height else 0
                out_w = self.size.out_width if self.size.out_width else 0
                out_h = self.size.out_height if self.size.out_height else 0
        else:
            in_w = self.size.in_width if self.size.in_width else 0
            in_h = self.size.in_height if self.size.in_height else 0
            out_w = self.size.out_width if self.size.out_width else 0
            out_h = self.size.out_height if self.size.out_height else 0

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
            in_shape, in_w, in_h, self.size.in_diameter)
        out_half_w, out_half_h = halves(
            out_shape, out_w, out_h, self.size.out_diameter)

        # Center displacement
        center_vec = outlet_pt - inlet_pt

        # Project onto global axes for consistent measurements
        # Vertical is always Z-axis (up/down)
        center_v = center_vec.Z

        # Horizontal: project onto XY plane and decompose into perpendicular directions
        # We need to measure offset relative to the duct's orientation, not flow direction
        # Use basis vectors if available to determine duct orientation
        if inlet_basis_x and inlet_basis_y:
            # Use actual duct orientation from connectors
            bx = XYZ(inlet_basis_x.X, inlet_basis_x.Y, inlet_basis_x.Z)
            by = XYZ(inlet_basis_y.X, inlet_basis_y.Y, inlet_basis_y.Z)

            # Project center_vec onto these basis directions
            # But map them to vertical/horizontal based on their alignment with global Z
            bx_z = abs(bx.Z)
            by_z = abs(by.Z)

            if bx_z > by_z and bx_z > 0.5:
                # basis_x is vertical (duct rotated 90°)
                center_h = center_vec.dot(by)
                center_v = center_vec.dot(bx)
                # Swap dimensions
                in_w = self.size.in_height if self.size.in_height else 0
                in_h = self.size.in_width if self.size.in_width else 0
                out_w = self.size.out_height if self.size.out_height else 0
                out_h = self.size.out_width if self.size.out_width else 0
            else:
                # Normal: basis_y is vertical (or more vertical)
                center_h = center_vec.dot(bx)
                center_v = center_vec.dot(by)
                in_w = self.size.in_width if self.size.in_width else 0
                in_h = self.size.in_height if self.size.in_height else 0
                out_w = self.size.out_width if self.size.out_width else 0
                out_h = self.size.out_height if self.size.out_height else 0
        else:
            # No basis vectors - use global coordinates
            # Vertical is Z
            center_v = center_vec.Z
            # Horizontal magnitude in XY plane
            center_h = (center_vec.X ** 2 + center_vec.Y ** 2) ** 0.5
            in_w = self.size.in_width if self.size.in_width else 0
            in_h = self.size.in_height if self.size.in_height else 0
            out_w = self.size.out_width if self.size.out_width else 0
            out_h = self.size.out_height if self.size.out_height else 0

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
