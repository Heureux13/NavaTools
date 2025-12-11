# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

# Imports
# =========================================================================
import math
from xyz import XYZ

# Class
# =========================================================================


class Points:
    def __init__(self, inlet, outlet, right=None, up=None):
        self.inlet = inlet
        self.outlet = outlet
        self.right = right
        self.up = up

    def round(self, diameter):
        radius = diameter / 2

        # Generate 360 points around the circle (one per degree)
        points = []
        for degree in range(360):
            rad = math.radians(degree)
            x = radius * math.cos(rad)
            y = radius * math.sin(rad)
            # Point on circle = center + (x * right + y * up)
            point = self.inlet + self.right * x + self.up * y
            points.append(point)

        return points

    def rectangle(self, width, height):
        half_w = width / 2
        half_h = height / 2

        # Rectangle corners (clockwise: top right, bottom right, bottom left, top left)
        corners = [
            self.inlet + self.right * half_w + self.up * half_h,   # Top right
            self.inlet + self.right * half_w +
            self.up * (-half_h),  # Bottom right
            self.inlet + self.right * (-half_w) +
            self.up * (-half_h),  # Bottom left
            self.inlet + self.right * (-half_w) + self.up * half_h,  # Top left
        ]

        # Generate 360 points: 90 per side
        points = []
        n_side = 90
        for i in range(4):
            start = corners[i]
            end = corners[(i+1) % 4]
            for j in range(n_side):
                t = j / n_side
                pt = XYZ(
                    start.X + t * (end.X - start.X),
                    start.Y + t * (end.Y - start.Y),
                    start.Z + t * (end.Z - start.Z)
                )
                points.append(pt)
        return points

    def oval(self, width, height):
        # True duct oval: two semicircular ends (height) and two straight sides
        minor_radius = height / 2
        straight_length = width - height

        points = []
        # Number of points for each semicircle and each straight (total 360 for smoothness)
        n_semi = 90  # points per semicircle (180 total)
        n_straight = 90  # points per straight (180 total)

        # First semicircle (left end, from bottom to top)
        for i in range(n_semi):
            theta = math.pi/2 + (i / (n_semi-1)) * math.pi  # from 90° to 270°
            x = -straight_length/2 + minor_radius * math.cos(theta)
            y = minor_radius * math.sin(theta)
            pt = self.inlet + self.right * x + self.up * y
            points.append(pt)

        # Top straight (from left to right)
        for i in range(1, n_straight):  # skip first point to avoid duplicate
            x = -straight_length/2 + minor_radius + \
                (i / (n_straight-1)) * (straight_length)
            y = minor_radius
            pt = self.inlet + self.right * x + self.up * y
            points.append(pt)

        # Second semicircle (right end, from top to bottom)
        for i in range(n_semi):
            theta = 3*math.pi/2 + (i / (n_semi-1)) * \
                math.pi  # from 270° to 450°
            x = straight_length/2 + minor_radius * math.cos(theta)
            y = minor_radius * math.sin(theta)
            pt = self.inlet + self.right * x + self.up * y
            points.append(pt)

        # Bottom straight (from right to left)
        for i in range(1, n_straight):  # skip first point to avoid duplicate
            x = straight_length/2 - minor_radius - \
                (i / (n_straight-1)) * (straight_length)
            y = -minor_radius
            pt = self.inlet + self.right * x + self.up * y
            points.append(pt)

        return points
