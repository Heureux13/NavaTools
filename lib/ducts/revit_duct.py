# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

# Standard library
# =========================================================
from ducts.revit_xyz import RevitXYZ
from geometry.size import Size
from geometry.offsets import Offsets
from Autodesk.Revit.DB import (
    ElementId,
    FilteredElementCollector,
    BuiltInCategory,
    UnitUtils,
    FabricationPart,
    UnitTypeId,
    ConnectorType
)
import re
import logging
import math
from enum import Enum
from ducts.connector_thresholds import (
    CONNECTOR_THRESHOLDS,
    DEFAULT_SHORT_THRESHOLD_IN,
    JointSize,
)

from config.parameters_registry import *

# Thrid Party
from pyrevit import DB, revit, script

#
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


def get_element_id_value(element_id):
    """Get integer value from ElementId, handling version differences."""
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue


# Material Density Class
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


# Duct Angle Allowance
# ====================================================
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


# Revut Duct Class
# ============================================================
class RevitDuct:
    def __init__(self, doc, view, element):
        self.doc = doc
        self.view = view
        self.element = element

    def get_connectors(self):
        """Return a list of all connectors for this duct element."""
        try:
            return list(self.element.ConnectorManager.Connectors)
        except Exception:
            return []

    @property
    def id(self):
        return self.element.Id.Value if self.element else None

    @property
    def category(self):
        return self.element.Category.Name if self.element and self.element.Category else None

    def get_connector(self, index):
        connectors = list(self.element.ConnectorManager.Connectors)
        if 0 <= index < len(connectors):
            return connectors[index]
        return None

    def _get_param(self, name, unit=None, as_type="string", required=False):
        p = self.element.LookupParameter(name)
        if not p:
            if required:
                raise KeyError(
                    "Missing parameter '{}' on element {}".format(
                        name,
                        self.element.Id,
                    ))
            return None

        try:
            if as_type == "double":
                val = p.AsDouble()
                if val is None:
                    return None
                if unit:
                    val = UnitUtils.ConvertFromInternalUnits(val, unit)
                return float(val)
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
            # convert any unexpected Revit exception into None to keep callers
            # deterministic
            return None

    def _inlet_outlet_from_revit_xyz(self):
        """Get inlet/outlet data via RevitXYZ (connector-based, no curve helper)."""
        xyz_extractor = RevitXYZ(self.element)
        inlet_data, outlet_data = xyz_extractor.inlet_outlet_data()
        if inlet_data and outlet_data:
            return inlet_data, outlet_data
        return None, None

    @property
    def size(self):
        return self._get_param(RVT_SIZE)

    @property
    def offset_top(self):
        inlet_data, outlet_data = self._inlet_outlet_from_revit_xyz()
        if not inlet_data or not outlet_data:
            return None

        offsets_obj = Offsets(inlet_data, outlet_data, Size(self.size))
        result = offsets_obj.calculate()
        return result['top'] if result else None

    @property
    def offset_bottom(self):
        inlet_data, outlet_data = self._inlet_outlet_from_revit_xyz()
        if not inlet_data or not outlet_data:
            return None

        offsets_obj = Offsets(inlet_data, outlet_data, Size(self.size))
        result = offsets_obj.calculate()
        return result['bottom'] if result else None

    @property
    def offset_left(self):
        inlet_data, outlet_data = self._inlet_outlet_from_revit_xyz()
        if not inlet_data or not outlet_data:
            return None

        offsets_obj = Offsets(inlet_data, outlet_data, Size(self.size))
        result = offsets_obj.calculate()
        return result['left'] if result else None

    @property
    def offset_right(self):
        inlet_data, outlet_data = self._inlet_outlet_from_revit_xyz()
        if not inlet_data or not outlet_data:
            return None

        offsets_obj = Offsets(inlet_data, outlet_data, Size(self.size))
        result = offsets_obj.calculate()
        return result['right'] if result else None

    @property
    def offset_center_h(self):
        inlet_data, outlet_data = self._inlet_outlet_from_revit_xyz()
        if not inlet_data or not outlet_data:
            return None

        offsets_obj = Offsets(inlet_data, outlet_data, Size(self.size))
        result = offsets_obj.calculate()
        return result['center_horizontal'] if result else None

    @property
    def offset_center_v(self):
        inlet_data, outlet_data = self._inlet_outlet_from_revit_xyz()
        if not inlet_data or not outlet_data:
            return None

        offsets_obj = Offsets(inlet_data, outlet_data, Size(self.size))
        result = offsets_obj.calculate()
        return result['center_vertical'] if result else None

    @property
    def centerline_length(self):
        return self._get_param(
            NDBS_CENTERLINE_LENGTH, unit=UnitTypeId.Inches, as_type="double")

    @property
    def length(self):
        result_0 = self._get_param(
            RVT_LENGTH, unit=UnitTypeId.Inches, as_type="double")
        if result_0 is not None:
            return result_0

        result_1 = self._get_param(
            NDBS_CENTERLINE_LENGTH, unit=UnitTypeId.Inches, as_type="double")
        if result_1 is not None:
            return result_1

        else:
            return None

    @property
    def size_in(self):
        size_str = self.size
        if size_str:
            size_obj = Size(size_str)
            if size_obj.in_size is not None:
                return size_obj.in_size

    @property
    def size_out(self):
        size_str = self.size
        if size_str:
            size_obj = Size(size_str)
            if size_obj.out_size is not None:
                return size_obj.out_size

    @property
    def diameter_in(self):
        size_str = self.size
        if size_str:
            size_obj = Size(size_str)
            if size_obj.in_diameter is not None:
                return size_obj.in_diameter

        return self._get_param(
            RVT_MAIN_PRIMARY_DIAMETER, unit=UnitTypeId.Inches, as_type="double")

    @property
    def diameter_out(self):
        size_str = self.size
        if size_str:
            size_obj = Size(size_str)
            if size_obj.out_diameter is not None:
                return size_obj.out_diameter

        return self._get_param(
            RVT_MAIN_SECONDARY_DIAMETER, unit=UnitTypeId.Inches, as_type="double")

    @property
    def height_in(self):
        size_str = self.size
        if size_str:
            size_obj = Size(size_str)
            if size_obj.in_height is not None:
                return size_obj.in_height

        return self._get_param(
            RVT_MAIN_PRIMARY_DEPTH, unit=UnitTypeId.Inches, as_type="double")

    @property
    def width_in(self):
        size_str = self.size
        if size_str:
            size_obj = Size(size_str)
            if size_obj.in_width is not None:
                return size_obj.in_width

        return self._get_param(
            RVT_MAIN_PRIMARY_WIDTH, unit=UnitTypeId.Inches, as_type="double")

    @property
    def width_out(self):
        size_str = self.size
        if size_str:
            size_obj = Size(size_str)
            if size_obj.out_width is not None:
                return size_obj.out_width

        return self._get_param(
            RVT_MAIN_SECONDARY_WIDTH, unit=UnitTypeId.Inches, as_type="double")

    @property
    def height_out(self):
        size_str = self.size
        if size_str:
            size_obj = Size(size_str)
            if size_obj.out_height is not None:
                return size_obj.out_height

        return self._get_param(
            RVT_MAIN_SECONDARY_DEPTH, unit=UnitTypeId.Inches, as_type="double")

    @property
    # Ex: TDF, S&D
    def connector_0_type(self):
        return self._get_param(NDBS_CONNECTOR0_END_CONDITION)

    @property
    # Ex: TDF, S&D
    def connector_1_type(self):
        return self._get_param(NDBS_CONNECTOR1_END_CONDITION)

    @property
    # Ex: TDF, S&D
    def connector_2_type(self):
        return self._get_param(NDBS_CONNECTOR2_END_CONDITION)

    @property
    def connector_0(self):
        return self.get_connector(0)

    @property
    def connector_1(self):
        return self.get_connector(1)

    @property
    def connector_2(self):
        return self.get_connector(2)

    def fully_connected(self):
        for connector in self.get_connectors():
            if connector.ConnectorType == ConnectorType.End:
                if not connector.IsConnected:
                    return False

        return True

    @property
    def extension_top(self):
        return self._get_param(
            NDBS_D_TOP_EXTENSION, unit=UnitTypeId.Inches, as_type="double")

    @property
    def extension_bottom(self):
        return self._get_param(
            NDBS_D_BOTTOM_EXTENSION, unit=UnitTypeId.Inches, as_type="double")

    @property
    def extension_right(self):
        return self._get_param(
            NDBS_D_RIGHT_EXTENSION, unit=UnitTypeId.Inches, as_type="double")

    @property
    def extension_left(self):
        return self._get_param(
            NDBS_D_LEFT_EXTENSION, unit=UnitTypeId.Inches, as_type="double")

    @property
    def duty(self):
        return self._get_param(RVT_SYSTEM_ABBREVIATION)

    @property
    def family(self):
        fam = self._get_param(RVT_FAMILY)
        if fam:
            return fam

        fam = self._get_param(NDBS_FAMILY)
        if fam:
            return fam

        return None

    @property
    def is_double_wall(self):
        return self._get_param(NDBS_HAS_DOUBLE_WALL)

    @property
    def has_insulation(self):
        return self._get_param(NDBS_HAS_INSULATION)

    @property
    def insulation_type(self):
        raw = self._get_param(RVT_INSULATION_SPECIFICATION)
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
        raw = self._get_param(RVT_INSULATION_SPECIFICATION)
        if not raw:
            # Use logger instead of print to avoid polluting pyRevit output
            log.debug(
                "Insulation Specification parameter not found or empty "
                "on element {}".format(self.id))
            return None

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
                "Sheet metal area parameter not found on element {}".format(
                    self.id))
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
            log.debug(
                "Weight parameter not found on element {}".format(
                    self.id))
            return None

        if not isinstance(insul_lb, (int, float)):
            insul_lb = 0.0

        return round(metal_lb + insul_lb, 2)

    @property
    def weight(self):
        return self._get_param(
            RVT_WEIGHT, unit=UnitTypeId.PoundsMass, as_type="double")

    @property
    def service(self):
        return self._get_param(NDBS_SERVICE_NAME)

    @property
    def inner_radius(self):
        return self._get_param(NDBS_D_INNER_RADIUS)

    @property
    def metal_area(self):
        return self._get_param(
            NDBS_SHEET_METAL_AREA, unit=UnitTypeId.SquareFeet, as_type="double")

    @property
    def angle(self):
        raw = self._get_param(RVT_ANGLE)
        if raw:
            cleaned = raw.replace(u"\xb0", "")
            try:
                return float(cleaned)
            except ValueError:
                return cleaned
        return None

    @property
    def joint_size(self):
        fam = (self.family or "").strip().lower()
        conn0 = (self.connector_0_type or "").strip().lower()
        conn1 = (self.connector_1_type or "").strip().lower()

        if conn0 != conn1:
            return JointSize.INVALID

        threshold = None
        for (k_family, k_conn), k_threshold in CONNECTOR_THRESHOLDS.items():
            if fam == (k_family or "").strip().lower() and conn0 == (k_conn or "").strip().lower():
                threshold = k_threshold
                break

        if threshold is None:
            if self.length is None:
                return JointSize.INVALID
            threshold = DEFAULT_SHORT_THRESHOLD_IN
        if self.length is None:
            return JointSize.INVALID

        # Apply a small tolerance to avoid classifying near-threshold parts as short
        tol = 0.01  # inches
        if self.length < (threshold - tol):
            return JointSize.SHORT
        if abs(self.length - threshold) <= tol:
            return JointSize.FULL
        if self.length > (threshold + tol):
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
    def from_selection(cls, uidoc, doc, view=None):
        sel_ids = uidoc.Selection.GetElementIds()
        if not sel_ids:
            return []

        elements = [doc.GetElement(elid) for elid in sel_ids]
        revit_year = get_revit_year(app)

        if revit_year <= 2023:
            duct = [
                el for el in elements if isinstance(
                    el, FabricationPart) and el.Category and el.Category.Id.IntegerValue == int(
                    BuiltInCategory.OST_FabricationDuctwork)]

        else:
            duct = [
                el for el in elements if isinstance(
                    el, FabricationPart) and el.Category and el.Category.Id.Value == int(
                    BuiltInCategory.OST_FabricationDuctwork)]

        return [cls(doc, view or uidoc.ActiveView, du) for du in duct]

    def get_connected_elements(self, connector_index=0):
        """Gets all elements connected to the selected element"""
        connector = self.get_connector(connector_index)
        connected_elements = []

        if connector and connector.IsConnected:
            for ref_conn in connector.AllRefs:
                if ref_conn.Owner.Id != self.element.Id:
                    connected_elements.append(ref_conn.Owner)
        return connected_elements
