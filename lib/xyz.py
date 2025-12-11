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

# Class
# =========================================================================


class XYZ:
    def __init__(self, x, y, z):
        self.X = x
        self.Y = y
        self.Z = z

    def __add__(self, other):
        return XYZ(self.X + other.X, self.Y + other.Y, self.Z + other.Z)

    def __sub__(self, other):
        return XYZ(self.X - other.X, self.Y - other.Y, self.Z - other.Z)

    def __mul__(self, scalar):
        return XYZ(self.X * scalar, self.Y * scalar, self.Z * scalar)

    def __truediv__(self, scalar):
        return XYZ(self.X / scalar, self.Y / scalar, self.Z / scalar)

    def dot(self, other):
        return self.X * other.X + self.Y * other.Y + self.Z * other.Z

    def cross(self, other):
        return XYZ(
            self.Y * other.Z - self.Z * other.Y,
            self.Z * other.X - self.X * other.Z,
            self.X * other.Y - self.Y * other.X,
        )

    def distance_to(self, other):
        return math.sqrt((self.X - other.X) ** 2 + (self.Y - other.Y) ** 2 + (self.Z - other.Z) ** 2)

    def __repr__(self):
        return "XYZ({}, {}, {})".format(self.X, self.Y, self.Z)

    def normalize(self):
        """Return a normalized (unit length) version of this vector"""
        direction = math.sqrt(self.X ** 2 + self.Y ** 2 + self.Z ** 2)
        if direction == 0:
            return XYZ(0, 0, 0)
        return XYZ(self.X / direction, self.Y / direction, self.Z / direction)
