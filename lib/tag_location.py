# tag_location.py
# Copyright (c) 2025 Jose Francisco Nava Perez
# All rights reserved. No part of this code may be reproduced without permission.

import math


class Point:
    def __init__(self, x, y, z=0):
        self.x, self.y, self.z = x, y, z

    def __add__(self, other):
        if isinstance(other, Vector):
            # point + vector = point
            return Point(self.x + other.x, self.y + other.y, self.z + other.z)
        elif isinstance(other, Point):
            # not usually defined, but you can allow it if you want
            return Point(self.x + other.x, self.y + other.y, self.z + other.z)
        raise TypeError("Unsupported addition with Point")

    def __sub__(self, other):
        if isinstance(other, Point):
            # point - point = vector
            return Vector(self.x - other.x, self.y - other.y, self.z - other.z)
        elif isinstance(other, Vector):
            # point - vector = point
            return Point(self.x - other.x, self.y - other.y, self.z - other.z)
        raise TypeError("Unsupported subtraction with Point")

    def __repr__(self):
        return "Point({0}, {1}, {2})".format(self.x, self.y, self.z)


class Line:
    def __init__(self, start, end):
        self.start = start
        self.end = end

    @property
    def direction(self):
        # returns a normalized vector from start to end
        return (self.end - self.start).normalize()

    def evaluate(self, t):
        # linear interpolation between start and end
        return Point(
            self.start.x + (self.end.x - self.start.x) * t,
            self.start.y + (self.end.y - self.start.y) * t,
            self.start.z + (self.end.z - self.start.z) * t
        )


class Vector:
    def __init__(self, x, y, z=0):
        self.x, self.y, self.z = x, y, z

    def normalize(self):
        length = math.sqrt(self.x**2 + self.y**2 + self.z**2)
        return Vector(self.x/length, self.y/length, self.z/length)

    def cross(self, other):
        return Vector(self.y*other.z - self.z*other.y,
                      self.z*other.x - self.x*other.z,
                      self.x*other.y - self.y*other.x)

    def dot(self, other):
        return self.x*other.x + self.y*other.y + self.z*other.z

    def angle_to(self, other):
        dot = self.dot(other)
        det = self.cross(other)
        return math.atan2(det.z, dot)

    def __mul__(self, scalar):
        return Vector(self.x*scalar, self.y*scalar, self.z*scalar)

    def __rmul__(self, scalar):
        return self.__mul__(scalar)

    def __add__(self, other):
        return Vector(self.x + other.x, self.y + other.y, self.z + other.z)

    def __sub__(self, other):
        return Vector(self.x - other.x, self.y - other.y, self.z - other.z)

    def __repr__(self):
        return "Vector({0}, {1}, {2})".format(self.x, self.y, self.z)


if __name__ == "__main__":
    # Straight line test
    line = Line(Point(0, 0), Point(10, 0))
    mid = line.evaluate(0.5)
    print("Midpoint:", mid)              # Expect Point(5, 0, 0)
    print("Direction:", line.direction)  # Expect Vector(1, 0, 0)

    # Angle test
    v1 = Vector(1, 0)
    v2 = Vector(0, 1)
    print("Angle between:", v1.angle_to(v2))  # Expect 1.5708

    # Offset test
    offset = v1.normalize() * 2
    print("Offset point:", mid + offset)      # Expect Point(7, 0, 0)
