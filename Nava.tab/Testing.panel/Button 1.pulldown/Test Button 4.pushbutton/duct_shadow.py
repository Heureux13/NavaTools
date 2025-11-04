# duct_shadow.py
# Copyright (c) 2025 Jose Francisco Nava Perez
# All rights reserved. No part of this code may be reproduced without permission.

from Autodesk.Revit.DB import XYZ


def _is_zero(v, eps=1e-9):
    return abs(v.X) < eps and abs(v.Y) < eps and abs(v.Z) < eps


def _safe_perp(a, hint):
    # returns a unit vector perpendicular to 'a', using 'hint' if possible
    c = a.CrossProduct(hint)
    if _is_zero(c):
        # pick a canonical axis not colinear with 'a'
        fallback = XYZ(1, 0, 0) if abs(a.X) < 0.9 else XYZ(0, 1, 0)
        c = a.CrossProduct(fallback)
    return c.Normalize()


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

        # Build a stable local frame around the duct axis
        # cross-section axis 1
        self.Yaxis = _safe_perp(self.axis_vec, self.view_dir)
        self.Zaxis = self.axis_vec.CrossProduct(
            self.Yaxis).Normalize()  # cross-section axis 2

        # Build openings
        self.start_corners = self._build_opening(self.start_pt)
        self.end_corners = self._build_opening(
            self.start_pt + self.axis_vec.Multiply(self.length)
        )
        self.all_corners = self.start_corners + self.end_corners
        self.shadow_pts = [self._project_to_view(p) for p in self.all_corners]

    def _build_opening(self, center):
        w2, h2 = self.width / 2.0, self.height / 2.0
        return [
            center + self.Yaxis.Multiply(w2) + self.Zaxis.Multiply(h2),
            center - self.Yaxis.Multiply(w2) + self.Zaxis.Multiply(h2),
            center - self.Yaxis.Multiply(w2) - self.Zaxis.Multiply(h2),
            center + self.Yaxis.Multiply(w2) - self.Zaxis.Multiply(h2)
        ]

    def _project_to_view(self, pt):
        rel = pt - self.origin
        return (rel.DotProduct(self.right), rel.DotProduct(self.up))

    def get_anchor(self, position="bottom_left", offset=(0, 0)):
        xs = [p[0] for p in self.shadow_pts]
        ys = [p[1] for p in self.shadow_pts]

        if position == "bottom_left":
            x, y = min(xs), min(ys)
        elif position == "top_left":
            x, y = min(xs), max(ys)
        elif position == "top_right":
            x, y = max(xs), max(ys)
        elif position == "bottom_right":
            x, y = max(xs), min(ys)
        else:  # centroid
            x, y = sum(xs) / len(xs), sum(ys) / len(ys)

        dx, dy = offset
        x += dx
        y += dy

        # Convert back into model space
        return self.origin + self.right.Multiply(x) + self.up.Multiply(y)
