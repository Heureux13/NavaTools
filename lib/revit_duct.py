# -*- coding: utf-8 -*-
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""

from Autodesk.Revit.DB import *
from pyrevit import revit, forms, DB
from Autodesk.Revit.UI import UIDocument
from Autodesk.Revit.ApplicationServices import Application
import re
from enum import Enum

# Variables
app   = __revit__.Application #type: Application
uidoc = __revit__.ActiveUIDocument #type: UIDocument
doc   = revit.doc #type: Document
view  = revit.active_view

# Define Constants
CONNECTOR_THRESHOLDS = {
    ("Straight", "TDC"): 56.25,
    ("Straight", "Standing S&D"): 59.0,
    ("Straight", "S&D"): 59.0,
    ("Tube", "AccuFlange"): 120.0,
    ("Tube", "GRC_Swage-Female"): 120.0,
}

class JointSize(Enum):
    FULL = "full"
    SHORT = "short"
    INVALID = "invalid"

class RevitDuct:
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
            print("[MISSING PARAM] '{0}' on element {1}".format(name, self.id))
            return None
        if as_type == "double":
            val = p.AsDouble()
            if unit:
                val = UnitUtils.ConvertFromInternalUnits(val, unit)
            return round(val, 2)
        if as_type == "int":
            return p.AsInteger()
        return p.AsString() or p.AsValueString()
    
    @property
    def size(self):
        return self._get_param("Size")
    
    @property
    def length(self):
        return self._get_param("NaviateDBS_D_Length", unit=UnitTypeId.Inches, as_type="double")
    
    @property
    def width(self):
        return self._get_param("NaviateDBS_D_Width", unit=UnitTypeId.Inches, as_type="double")

    @property
    def depth(self):
        return self._get_param("NaviateDBS_D_Depth", unit=UnitTypeId.Inches, as_type="double")
    
    @property
    def connector_0(self):
        return self._get_param("NaviateDBS_Connector0_EndCondition")

    @property
    def connector_1(self):
        return self._get_param("NaviateDBS_Connector1_EndCondition")
    
    @property
    def connector_2(self):
        return self._get_param("NaviateDBS_Connector2_EndCondition")

    @property
    def system_abbreviation(self):
        return self._get_param("System Abbreviation")

    @property
    def family(self):
        return self._get_param("NaviateDBS_Family")
    
    @property
    def double_wall(self):
        return self._get_param("NaviateDBS_HasDoubleWall")
    
    @property
    def insulation(self):
        return self._get_param("NaviateDBS_HasInsulation")
    
    @property
    def insulation_specification(self):
        return self._get_param("Insulation Specification")
    
    @property
    def service(self):
        return self._get_param("NaviateDBS_ServiceName")
    
    @property
    def inner_radius(self):
        return self._get_param("NaviateDBS_InnerRadius")
    
    @property
    def extension_top(self):
        return self._get_param("NaviateDBS_D_Top Extension", unit=UnitTypeId.Inches, as_type="double")
    
    @property
    def extension_bottom(self):
        return self._get_param("NaviateDBS_D_Bottom Extension", unit=UnitTypeId.Inches, as_type="double")
    
    @property
    def extension_right(self):
        return self._get_param("NaviateDBS_D_Right Extension", unit=UnitTypeId.Inches, as_type="double")
    
    @property
    def extension_left(self):
        return self._get_param("NaviateDBS_D_Left Extension", unit=UnitTypeId.Inches, as_type="double")
    
    @property
    def area(self):
        return self._get_param("NaviateDBS_SheetMetalArea", unit=UnitTypeId.SquareFeet, as_type="double")
    
    @property
    def weight(self):
        return self._get_param("NaviateDBS_Weight", as_type="double")
    
    @property
    def total_weight(self):
        base = self.weight or 0.0
        liner = getattr(self, "liner_weight", 0.0) or 0.0
        insulation = getattr(self, "insulation_weight", 0.0) or 0.0
        return round(base + liner + insulation, 2)
    
    @property
    def angle(self):
        raw = self._get_param("Angle")
        if raw:
            cleaned = raw.replace(u"\xb0", "")
            try:
                return float(cleaned)
            except ValueError:
                return cleaned
        return None

    @property
    # Checks to see if a straight joint is full or short size.
    def is_full_joint(self):
        conn0 = (self.connector_0 or "").strip()
        conn1 = (self.connector_1 or "").strip()
        key = (self.family, conn0)

        if conn0 != conn1:
            return JointSize.INVALID

        threshold = CONNECTOR_THRESHOLDS.get(key)
        if threshold is None or self.length is None:
            return JointSize.INVALID

        return JointSize.SHORT if self.length < threshold else JointSize.FULL


    def get_selected_elements(doc, uidoc, filter_types=None, categories=None, from_selection=True):
        if from_selection:
            sel_ids = uidoc.Selection.GetElementIds()
            elements = [doc.GetElement(eid) for eid in sel_ids]
        else:
            collector = DB.FilteredElementCollector(doc).WhereElementIsNotElementType()
            if categories:
                collector = collector.OfCategory(categories[0])

                elements = list(collector)

        if filter_types:
            elements = [el for el in elements if isinstance(el, filter_types)]
        return elements