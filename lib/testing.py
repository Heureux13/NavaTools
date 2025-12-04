# Simple Box class for 3D geometry
import math


class Box:
    def __init__(self, min_xyz, max_xyz):
        self.min = min_xyz  # XYZ: lower corner
        self.max = max_xyz  # XYZ: upper corner

    def top(self):
        return self.max.Z

    def bottom(self):
        return self.min.Z

    def left(self):
        return min(self.min.X, self.max.X)

    def right(self):
        return max(self.min.X, self.max.X)

    def front(self):
        return min(self.min.Y, self.max.Y)

    def back(self):
        return max(self.min.Y, self.max.Y)

    def height(self):
        return abs(self.max.Z - self.min.Z)

    def width(self):
        return abs(self.max.X - self.min.X)

    def depth(self):
        return abs(self.max.Y - self.min.Y)

    def angle_to_floor(self):
        # Returns angle (degrees) between box's vertical axis and Z axis (floor)
        # For a box aligned to world axes, this is always 0
        # For a duct, you might use the vector from min to max
        v = self.max - self.min
        vertical = XYZ(0, 0, 1)
        dot = v.dot(vertical)
        mag_v = math.sqrt(v.X**2 + v.Y**2 + v.Z**2)
        mag_vert = 1.0
        cos_theta = dot / (mag_v * mag_vert) if mag_v != 0 else 0
        angle_rad = math.acos(max(-1, min(1, cos_theta)))
        return math.degrees(angle_rad)

    def __repr__(self):
        return f"Box(min={self.min}, max={self.max})"


"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

# Imports
# ===========================================================================


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
            self.X * other.Y - self.Y * other.X
        )

    def distance_to(self, other):
        return math.sqrt((self.X - other.X) ** 2 + (self.Y - other.Y) ** 2 + (self.Z - other.Z) ** 2)

    def __repr__(self):
        return f"XYZ({self.X}, {self.Y}, {self.Z})"


if __name__ == "__main__":
    # Example usage of XYZ for 3D math
    a = XYZ(1, 2, 3)
    b = XYZ(4, 5, 6)
    print(f"{a} is vector a")
    print(f"{b} is vector b")
    print(f"{a + b} is a + b")
    print(f"{a - b} is a - b")
    print(f"{a.dot(b)} is a . b (dot product)")
    print(f"{a.cross(b)} is a x b (cross product)")
    print(f"{a.distance_to(b)} is the distance from a to b")
    # You can still use XYZ for mock curve endpoints as before

    class MockCurve(object):
        def __init__(self, start, end):
            self.start = XYZ(*start)
            self.end = XYZ(*end)

        def GetEndPoint(self, idx):
            return self.start if idx == 0 else self.end

        def Evaluate(self, t, _):
            x = self.start.X + t * (self.end.X - self.start.X)
            y = self.start.Y + t * (self.end.Y - self.start.Y)
            z = self.start.Z + t * (self.end.Z - self.start.Z)
            return XYZ(x, y, z)
    start = (8.115217245, -20.102905699, 10.500000000)
    end = (8.115217245, -12.181281631, 10.500000000)
    curve = MockCurve(start, end)
    print(f"{curve.GetEndPoint(0)} is curve start")
    print(f"{curve.GetEndPoint(1)} is curve end")
    print(f"{curve.Evaluate(0.5, True)} is curve midpoint")
