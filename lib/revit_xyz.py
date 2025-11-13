# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import UnitTypeId
from pyrevit import revit, script, forms, DB
from Autodesk.Revit.UI import UIDocument
from Autodesk.Revit.ApplicationServices import Application
from enum import Enum
import logging
import math
import re

# Variables
# ==================================================s   
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
view = revit.active_view
output = script.get_output()

# Logging
log = logging.getLogger("RevitDuct")

# Class logic
# ==================================================
class RevitXYZ(object):
    def __init__(self, element):
        self.element = element
        self.loc = getattr(element, "Location", None)
        self.curve = getattr(self.loc, "Curve", None) if self.loc else None
        self.doc = revit.doc
        self.view = revit.active_view

    def start_point(self):
        if self.curve:
            return self.curve.GetEndPoint(0)
        return None

    def end_point(self):
        if self.curve:
            return self.curve.GetEndPoint(1)
        return None

    def mid_point(self):
        if self.curve:
            return self.curve.Evaluate(0.5, True)
        return None

    def point_at(self, param=0.25):
        if self.curve:
            t = max(0.0, min(1.0, float(param)))
            return self.curve.Evaluate(t, True)
        return None

    def straight_joint_degree(self):
        """Returns the angle in degrees between the duct and the horizontal (XY) plane."""
        start = self.start_point()
        end = self.end_point()
        if not start or not end:
            return None

        dx = end.X - start.X
        dy = end.Y - start.Y
        dz = end.Z - start.Z

        horizontal_length = math.sqrt(dx**2 + dy**2)
        if horizontal_length == 0:
            return 90.0 if dz != 0 else 0.0

        angle_rad = math.atan2(dz, horizontal_length)
        angle_deg = math.degrees(angle_rad)
        return round(angle_deg, 2)

    def true_length(self):
        """Returns the true 3D length of the duct."""
        start = self.start_point()
        end = self.end_point()
        if not start or not end:
            return None

        dx = end.X - start.X
        dy = end.Y - start.Y
        dz = end.Z - start.Z

        length = math.sqrt(dx**2 + dy**2 + dz**2)
        return round(length, 2)