# -*- coding: utf-8 -*-
"""
=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
=========================================================================
"""

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import UnitTypeId
from pyrevit import revit, forms, DB
from Autodesk.Revit.UI import UIDocument
from Autodesk.Revit.ApplicationServices import Application
import re
from enum import Enum
import math

# Variables
app   = __revit__.Application #type: Application
uidoc = __revit__.ActiveUIDocument #type: UIDocument
doc   = revit.doc #type: Document
view  = revit.active_view

from Autodesk.Revit.DB import *
from pyrevit import revit

doc   = revit.doc
view  = revit.active_view

class RevitPipe:
    def __init__(self, doc, view, element):
        self.doc = doc
        self.view = view
        self.element = element

    @property
    def id(self):
        return self.element.Id.Value if self.element else None

    @property
    def category(self):
        return self.element.Category.Name if self.element and self.element.Category else None

    def _get_param(self, name, unit=None, as_type="string"):
        p = self.element.LookupParameter(name)
        if not p:
            return None
        if as_type == "double":
            val = p.AsDouble()
            if unit:
                val = UnitUtils.ConvertFromInternalUnits(val, unit)
            return round(val, 2)
        if as_type == "int":
            return p.AsInteger()
        return p.AsString() or p.AsValueString()

    # --- Pipe-specific properties ---
    @property
    def diameter(self):
        return self._get_param("Diameter", unit=UnitTypeId.Inches, as_type="double")

    @property
    def system_type(self):
        return self._get_param("Fabrication Service")

    @property
    def length(self):
        return self._get_param("Length", unit=UnitTypeId.Inches, as_type="double")

    @property
    def insulation(self):
        return self._get_param("Insulation Thickness", unit=UnitTypeId.Inches, as_type="double")

    @property
    def bop(self):
        return self._get_param("NaviateDBS_BottomOfPartElevation")

    @property
    def top(self):
        return self._get_param("NaviateDBS_TopOfPartElevation")

    @property
    def conn_0(self):
        return self._get_param("NaviateDBS_Connector0_EndCondition")

    @property
    def con_1(self):
        return self._get_param("NaviateDBS_Connector1_EndCondition")

    @property
    def family(self):
        return self._get_param("NaviateDBS_Family")

    @property
    def free_size(self):
        return self._get_param("NaviateDBS_FreeSize", unit=UnitTypeId.Inches, as_type="double")

    @property
    def service(self):
        return self._get_param("NaviateDBS_ServiceAbbreviation")

    @property
    def dry_weight(self):
        return self._get_param("NaviateDBS_Weight", as_type="double")

    @property
    def wet_weight(self):
        h = self.length
        r = self.free_size/2
        dw = self.dry_weight
        if None in (h, r, dw):
            return None
        v = math.pi*r*h
        v_sf = v/1728
        v_w = v_sf*62.4
        return round(dw + v_w, 2)


    # --- Class methods ---
    @classmethod
    def all(cls, doc, view=None):
        elements = (FilteredElementCollector(doc, view.Id if view else None)
                    .OfCategory(BuiltInCategory.OST_PipeCurves)
                    .WhereElementIsNotElementType()
                    .ToElements())
        return [cls(doc, view, el) for el in elements]

    @classmethod
    def count(cls, doc, view=None):
        return len(cls.all(doc, view))

    @classmethod
    def by_system_type(cls, doc, view, system_type_name):
        return [p for p in cls.all(doc, view)
                if p.system_type == system_type_name]