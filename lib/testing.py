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
            self.X * other.Y - self.Y * other.X
        )

    def distance_to(self, other):
        return math.sqrt((self.X - other.X) ** 2 + (self.Y - other.Y) ** 2 + (self.Z - other.Z) ** 2)

    def __repr__(self):
        return f"XYZ({self.X}, {self.Y}, {self.Z})"

    def round(self, diameter, outlet):
        radius = diameter / 2
        forward = (outlet - self).normalize()

        if abs(forward.Z) < 0.9:
            global_up = XYZ(0, 0, 1)  # World Z
        else:
            global_up = XYZ(1, 0, 0)  # World X

        right = forward.cross(global_up).normalize()
        up = right.cross(forward).normalize()

        # Generate 360 points around the circle
        points = []
        for degree in range(72):
            rad = math.radians(degree)
            x = radius * math.cos(rad)
            y = radius * math.sin(rad)
            # Point on circle = center + (x * right + y * forward)
            point = self + right * x + up * y
            points.append(point)

        return points

    def rectangle(self, width, height, outlet):
        forward = (outlet - self).normalize()

        if abs(forward.Z) < 0.9:
            global_up = XYZ(0, 0, 1)
        else:
            global_up = XYZ(1, 0, 0)

        right = forward.cross(global_up).normalize()
        up = right.cross(forward).normalize()

        half_w = width / 2
        half_h = height / 2

        corners = [
            self + right * half_w + up * half_h,   # Top right
            self + right * half_w + up * (-half_h),  # Bottom right
            self + right * (-half_w) + up * (-half_h),  # Bottom left
            self + right * (-half_w) + up * half_h,  # Top left
        ]
        return corners

    def oval(self, major_diameter, minor_diameter, outlet):
        major_radius = major_diameter / 2
        minor_radius = minor_diameter / 2

        forward = (outlet - self).normalize()

        if abs(forward.Z) < 0.9:
            global_up = XYZ(0, 0, 1)
        else:
            global_up = XYZ(1, 0, 0)

        right = forward.cross(global_up).normalize()
        up = right.cross(forward).normalize()

        # Generate 360 points around the ellipse
        points = []
        for degree in range(360):
            rad = math.radians(degree)
            x = major_radius * math.cos(rad)
            y = minor_radius * math.sin(rad)
            point = self + right * x + up * y
            points.append(point)

        return points

    def normalize(self):
        """Return a normalized (unit length) version of this vector"""
        direction = math.sqrt(self.X ** 2 + self.Y ** 2 + self.Z ** 2)
        if direction == 0:
            return XYZ(0, 0, 0)
        return XYZ(self.X / direction, self.Y / direction, self.Z / direction)


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

    # Test shape methods
    print("\n--- Testing Round Duct (Horizontal) ---")
    inlet = XYZ(0, 0, 10)
    outlet = XYZ(0, 10, 10)  # Horizontal duct going in Y direction
    circle_points = inlet.round(12, outlet)
    print(f"Generated {len(circle_points)} points for 12\" round duct")
    print(f"First point: {circle_points[0]}")
    print(f"90 degree point: {circle_points[90]}")

    print("\n--- Testing Round Duct (Vertical) ---")
    inlet_v = XYZ(0, 0, 10)
    outlet_v = XYZ(0, 0, 20)  # Vertical duct going up
    circle_points_v = inlet_v.round(12, outlet_v)
    print(
        f"Generated {len(circle_points_v)} points for vertical 12\" round duct")
    print(f"First point: {circle_points_v[0]}")
    print(f"90 degree point: {circle_points_v[90]}")

    print("\n--- Testing Rectangle Duct ---")
    rect_corners = inlet.rectangle(12, 8, outlet)
    print(f"Rectangle corners: {len(rect_corners)} points")
    for i, corner in enumerate(rect_corners):
        print(f"  Corner {i}: {corner}")

    print("\n--- Testing Oval Duct ---")
    oval_points = inlet.oval(20, 10, outlet)
    print(f"Generated {len(oval_points)} points for 20x10 oval duct")
    print(f"First point (major axis): {oval_points[0]}")
    print(f"90 degree point (minor axis): {oval_points[90]}")

    print("\n--- Testing normalize() ---")
    vector = XYZ(3, 4, 0)
    print(f"Original vector: {vector}")
    print(f"Length: {math.sqrt(vector.X**2 + vector.Y**2 + vector.Z**2)}")
    normalized = vector.normalize()
    print(f"Normalized vector: {normalized}")
    print(
        f"Normalized length: {math.sqrt(normalized.X**2 + normalized.Y**2 + normalized.Z**2)}")

    direction = outlet - inlet  # Vector from inlet to outlet
    print(f"\nDirection vector (outlet - inlet): {direction}")
    print(
        f"Direction length: {math.sqrt(direction.X**2 + direction.Y**2 + direction.Z**2)}")
    normalized_dir = direction.normalize()
    print(f"Normalized direction: {normalized_dir}")
    print(
        f"Normalized direction length: {math.sqrt(normalized_dir.X**2 + normalized_dir.Y**2 + normalized_dir.Z**2)}")
