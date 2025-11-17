# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

import re
import math
import logging
from enum import Enum
from pyrevit import DB, revit, script, forms
from Autodesk.Revit.ApplicationServices import Application
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import UnitTypeId
from Autodesk.Revit.UI import UIDocument
from Autodesk.Revit.DB import FabricationPart
import clr
clr.AddReference("RevitAPI")

# Variables
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
view = revit.active_view
output = script.get_output()

# Logging
log = logging.getLogger("RevitDuct")

# Define Constants
CONNECTOR_THRESHOLDS = {
    ("Straight", "TDC"): 56.00,
    ("Straight", "TDF"): 56.00,
    ("Straight", "Standing S&D"): 59.00,
    ("Straight", "Slip & Drive"): 59.00,
    ("Straight", "S&D"): 59.00,
    ("Tube", "AccuFlange"): 120.00,
    ("Tube", "GRC_Swage-Female"): 120.00,
    ("Spiral Duct", "Raw"): 120.00,
    ("Spiral Pipe", "Raw"): 120.00,
}

# Helpers
# ==================================================


def get_revit_year(app):
    name = app.VersionName
    for n in name.split():
        if n.isdigit():
            return int(n)
    return None


def is_plan_view(view):
    # Check if the view is a floor plan
    return view.ViewType == DB.ViewType.FloorPlan


def is_section_view(view):
    # Check if the view is a section
    return view.ViewType == DB.ViewType.Section


# Classes
# ==================================================
class MaterialDensity(Enum):
    LINER = (1.5, "lb/ft³", "Acoustic Liner")
    WRAP = (1.5, "lb/ft³", "Insulation Wrap")

    @property
    def density(self):
        return self.value[0]

    @property
    def unit(self):
        return self.value[1]

    @property
    def descrs(self):
        return self.value[2]


class JointSize(Enum):
    SHORT = "short"
    FULL = "full"
    LONG = "long"
    INVALID = "invalid"


class DuctAngleAllowance(Enum):
    HORIZONTAL = (0, 15)      # 0-15 degrees from horizontal
    # 75-90 degrees from horizontal (i.e., near vertical)
    VERTICAL = (75, 90)
    ANGLED = (16, 74)         # 16-74 degrees (neither horizontal nor vertical)

    @property
    def min_deg(self):
        return self.value[0]

    @property
    def max_deg(self):
        return self.value[1]

    def contains(self, angle):
        """Check if the given angle falls within this allowance."""
        return self.min_deg <= abs(angle) <= self.max_deg


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

    def _get_param(self, name, unit=None, as_type="string", required=False):
        # helper gettin parameters from revit
        p = self.element.LookupParameter(name)
        if not p:
            if required:
                raise KeyError(
                    "Missing parameter '{}' on element {}".format(name, self.element.Id))
            return None

        try:
            if as_type == "double":
                val = p.AsDouble()
                if val is None:
                    return None
                if unit:
                    val = UnitUtils.ConvertFromInternalUnits(val, unit)
                return float(round(val, 2))
            if as_type == "int":
                return p.AsInteger()
            if as_type == "elementid":
                eid = p.AsElementId()
                return eid if isinstance(eid, ElementId) else None
            # fallback string: prefer AsString, then AsValueString
            s = p.AsString()
            if s is None:
                s = p.AsValueString()
            return s
        except Exception:
            # convert any unexpected Revit exception into None to keep callers deterministic
            return None

    @property
    def size(self):
        return self._get_param("Size")

    @property
    def length(self):
        return self._get_param("Length", unit=UnitTypeId.Inches, as_type="double")

    @property
    def width(self):
        return self._get_param("Main Primary Width", unit=UnitTypeId.Inches, as_type="double")

    @property
    def depth(self):
        return self._get_param("Main Primary Depth", unit=UnitTypeId.Inches, as_type="double")

    @property
    # Accessed throught a paid version or an extended API
    def connector_0(self):
        return self._get_param("NaviateDBS_Connector0_EndCondition")

    @property
    # Accessed throught a paid version or an extended API
    def connector_1(self):
        return self._get_param("NaviateDBS_Connector1_EndCondition")

    @property
    # Accessed throught a paid version or an extended API
    def connector_2(self):
        return self._get_param("NaviateDBS_Connector2_EndCondition")

    # @property
    # # Accessed throught a paid version or an extended API
    # def connector_0_length(self):
    #     return self._get_param("NaviateDBS_Bottom Extension", unit=UnitTypeId.Inches, as_type="double")

    # @property
    # # Accessed throught a paid version or an extended API
    # def connector_1_length(self):
    #     return self._get_param("NaviateDBS_Top Extension", unit=UnitTypeId.Inches, as_type="double")

    # @property
    # # Accessed throught a paid version or an extended API
    # def connector_2_length(self):
    #     return self._get_param("NaviateDBS_left Extension", unit=UnitTypeId.Inches, as_type="double")

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
    def duty(self):
        return self._get_param("System Abbreviation")

    @property
    def offset_width(self):
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
    def insulation_type(self):
        raw = self._get_param("Insulation Specification")
        if not raw or not isinstance(raw, str):
            return MaterialDensity.LINER

        text = raw.lower()

        if re.search(r"\bliner\b", text):
            return MaterialDensity.LINER

        elif re.search(r"\binsulation\b", text):
            return MaterialDensity.WRAP

        else:
            return MaterialDensity.LINER

    @property
    def insulation_thickness(self):
        raw = self._get_param("Insulation Specification")
        if not raw:
            # Use logger instead of print to avoid polluting pyRevit output
            log.debug(
                "Insulation Specification parameter not found or empty on element {}".format(self.id))
            return None

        # Normalise various unicode quotation marks likely to appear in insulation specs
        # Original intent: convert smart inch and quote characters to plain ASCII for regex parsing
        cleaned = (raw
                   .replace(u"″", '"')   # double prime
                   .replace(u"”", '"')   # right double quotation mark
                   .replace(u"’", "'")  # right single quotation / apostrophe
                   )

        match = re.search(r"([\d\.]+)", cleaned)
        if match:
            try:
                return float(match.group(1))
            except ValueError:
                return None
        return None

    @property
    def weight_insulation(self):
        thic_in = self.insulation_thickness or 0.0
        area_ft2 = self.metal_area

        if area_ft2 is None:
            log.debug(
                "Sheet metal area parameter not found on element {}".format(self.id))
            return None

        material = self.insulation_type
        density_pcf = material.density

        weight_lb = density_pcf * (thic_in / 12) * area_ft2
        return round(weight_lb, 2)

    @property
    def weight_total(self):
        metal_lb = self.weight_metal
        insul_lb = self.weight_insulation

        if metal_lb is None:
            log.debug("Weight parameter not found on element {}".format(self.id))
            return None

        if not isinstance(insul_lb, (int, float)):
            insul_lb = 0.0

        return round(metal_lb + insul_lb, 2)

    @property
    def weight_metal(self):
        return self._get_param("Weight", unit=UnitTypeId.PoundsMass, as_type="double")

    @property
    def service(self):
        return self._get_param("NaviateDBS_ServiceName")

    @property
    def inner_radius(self):
        return self._get_param("NaviateDBS_InnerRadius")

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
        revit_year = get_revit_year(app)

        if revit_year <= 2023:
            duct = [
                el for el in elements if isinstance(el, FabricationPart)
                and el.Category
                and el.Category.Id.IntegerValue == int(BuiltInCategory.OST_FabricationDuctwork)
            ]

        else:
            duct = [
                el for el in elements if isinstance(el, FabricationPart)
                and el.Category
                and el.Category.Id.Value == int(BuiltInCategory.OST_FabricationDuctwork)
            ]

        return [cls(doc, view or uidoc.ActiveView, du) for du in duct]
