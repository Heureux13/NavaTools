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
    SHORT   = "short"
    FULL    = "full"
    LONG    = "long"
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
            forms.alert("[MISSING PARAM] '{0}' on element {1}".format(name, self.id))
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
    def duty(self):
        return self._get_param("System Abbreviation")

    @property
    def family(self):
        return self._get_param("NaviateDBS_Family")
    
    @property
    def is_double_wall(self):
        return self._get_param("NaviateDBS_HasDoubleWall")
    
    @property
    def has_insulation(self):
        return self._get_param("NaviateDBS_HasInsulation")

    @property
    def insulation(self):
        raw = self._get_param("Insulation Specification")
        if not raw:
            forms.alert("its not raw")
            return None

        cleaned = raw.replace("″", '"').replace("”", '"').replace("’", "'")

        match = re.search(r"([\d\.]+)", cleaned)
        if match:
            try:
                return float(match.group(1))
            except ValueError:
                return None
        return None
    
    @property
    def weight_insulation(self):
        thic_in = self.insulation if self.insulation is not None else 0.00
        area_ft2 = self.metal_area

        if thic_in is None or area_ft2 is None:
            return None
        
        density_pcf = 2.5

        thic_ft = thic_in / 12.0
        weight_lb = density_pcf * thic_ft * area_ft2
        return weight_lb
    
    @property
    def weight_total(self):
        metal_lb = self.weight_metal
        insul_lb = self.weight_insulation

        if metal_lb is None:
            forms.alert("No metal weight")
            return None
        
        return round(metal_lb + insul_lb, 2)
    
    @property
    def weight_metal(self):
        return self._get_param("NaviateDBS_Weight", unit=UnitTypeId.PoundsMass, as_type="double")
    
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
    def metal_area(self):
        return self._get_param("NaviateDBS_SheetMetalArea", unit=UnitTypeId.SquareFeet, as_type="double")
    
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
    # returns a four option varience, one being an error. these sizes and connections can bechanged easealy across various fabs
    def joint_size(self):
        conn0 = (self.connector_0 or "").strip()
        conn1 = (self.connector_1 or "").strip()
        key = (self.family, conn0)

        if conn0 != conn1:
            return JointSize.INVALID

        threshold = CONNECTOR_THRESHOLDS.get(key)
        if threshold is None or self.length is None:
            return JointSize.INVALID

        if self.length < threshold:
            return JointSize.SHORT
        if self.length == threshold:
            return JointSize.FULL
        if self.length > threshold:
            return JointSize.LONG

    @classmethod
    def all(cls, doc, view=None):
        """Return all duct elements wrapped as RevitDuct objects."""
        elements = (FilteredElementCollector(doc, view.Id if view else None)
                    .OfCategory(BuiltInCategory.OST_FabricationDuctwork)
                    .WhereElementIsNotElementType()
                    .ToElements())
        return [cls(doc, view, el) for el in elements]

    @classmethod
    def count(cls, doc, view=None):
        return len(cls.all(doc, view))

    @classmethod
    def by_system_type(cls, doc, view, system_type_name):
        return [d for d in cls.all(doc, view)
                if d.element.LookupParameter("System Type").AsString() == system_type_name]
    
    @classmethod
    def from_selection(cls, uidoc, doc, view=None):
        sel_ids = uidoc.Selection.GetElementIds()
        if not sel_ids:
            return []

        elements = [doc.GetElement(elid) for elid in sel_ids]

        duct = [
            el for el in elements if isinstance(el, FabricationPart)
                and el.Category
                and el.Category.Id.IntegerValue == int(BuiltInCategory.OST_FabricationDuctwork)
                ]
        return [cls(doc, view or uidoc.ActiveView, du) for du in duct]