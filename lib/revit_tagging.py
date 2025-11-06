# -*- coding: utf-8 -*-
############################################################################
# Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.
#
# This code and associated documentation files may not be copied, modified,
# distributed, or used in any form without the prior written permission of 
# the copyright holder.
############################################################################

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import UnitTypeId
from pyrevit import revit, forms, DB
from Autodesk.Revit.UI import UIDocument
from Autodesk.Revit.ApplicationServices import Application
from enum import Enum
import re

# Variables
# =======================================================================
app   = __revit__.Application           #type: Application
uidoc = __revit__.ActiveUIDocument      #type: UIDocument
doc   = revit.doc                       #type: Document
view  = revit.active_viewd

#Class
# =======================================================================
class RevitXYZ(object):
    def __init__(self, element):
        self.element = element
        self.loc = getattr(element, "Location", None)
        self.curve = getattr(self.loc, "Curve", None) if self.loc else None

    def start_point(self):
        if self.curve:
            return self.curve.GetEndPoint(0)
        return None

    def end_point(self):
        if self.curve:
            return self.curve.GetEndPoint(1)
        return None

    def midpoint(self):
        if self.curve:
            return self.curve.Evaluate(0.5, True)
        return None

    def point_at(self, param=0.25):
        #Get a point along the curve at a normalized parameter (0-1).
        if self.curve:
            return self.curve.Evaluate(param, True)
        return None