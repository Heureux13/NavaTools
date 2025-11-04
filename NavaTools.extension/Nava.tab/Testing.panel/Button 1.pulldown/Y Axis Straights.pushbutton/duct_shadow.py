# duct_shadow.py
# Copyright (c) 2025 Jose Francisco Nava Perez
# All rights reserved. No part of this code may be reproduced without permission.

from Autodesk.Revit.DB import XYZ


class DuctShadow:
    def __init__(self, start_pt, axis_vec, width, height, length, view):
        self.start_pt = start_pt
        self.axis_vec = axis_vec.Normalize()
        self.width = width
        self.height = height
        self.length = length
        self.view = view

        self.right = view.RightDirection.Normalize()
        self.up = view.UpDirection.Normalize()
        self.origin = view.Origin
        self.view_dir = view.ViewDirection.Normalize()

        self.Yaxis = self.view_dir.CrossProduct(self.axis_vec).Normalize()
        self.Zaxis = self.axis_vec.CrossProduct(self.Yaxis).Normalize()

        self.start_corners = self._build_opening(self.start_pt)
        self.end_corners = self._build_opening(
            self.start_pt + self.axis_vec * self.length)
        self.all_corners = self.start_corners + self.end_corners
        self.shadow_pts = [self._project_to_view(p) for p in self.all_corners]
        self.bottom_left = self._find_bottom_left()

    def _build_opening(self, center):
        w2, h2 = self.width / 2.0, self.height / 2.0
        return [
            center + (self.Yaxis * w2) + (self.Zaxis * h2),  # top right
            center - (self.Yaxis * w2) + (self.Zaxis * h2),  # top left
            center - (self.Yaxis * w2) - (self.Zaxis * h2),  # bottom left
            center + (self.Yaxis * w2) - (self.Zaxis * h2)   # bottom right
        ]

    def _project_to_view(self, pt):
        rel = pt - self.origin
        # 2D in view plane
        return (rel.DotProduct(self.right), rel.DotProduct(self.up))

    def _find_bottom_left(self):
        minX = min(p[0] for p in self.shadow_pts)
        minY = min(p[1] for p in self.shadow_pts)
        return XYZ(minX, minY, 0)

    def print_debug(self):
        print("Start Corners:")
        for pt in self.start_corners:
            print("  ", pt)
        print("End Corners:")
        for pt in self.end_corners:
            print("  ", pt)
        print("Projected Shadow Points:")
        for pt in self.shadow_pts:
            print("  (%.3f, %.3f)" % pt)
        print("Bottom-Left Anchor for Tag:")
        print("  ", self.bottom_left)
