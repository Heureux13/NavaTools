# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

# Standard library
# =========================================================
from revit_xyz import RevitXYZ
from size import Size
from offsets import Offsets
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


# Joint Size Class
# ====================================================
class JointSize(Enum):
    SHORT = "short"
    FULL = "full"
    LONG = "long"
    INVALID = "invalid"


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
        return self._get_param("Size")

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
            "NaviateDBS_CenterlineLength", unit=UnitTypeId.Inches, as_type="double")

    @property
    def length(self):
        result_0 = self._get_param(
            "Length", unit=UnitTypeId.Inches, as_type="double")
        if result_0 is not None:
            return result_0

        result_1 = self._get_param(
            "NaviateDBS_CenterlineLength", unit=UnitTypeId.Inches, as_type="double")
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
            "Main Primary Diameter", unit=UnitTypeId.Inches, as_type="double")

    @property
    def diameter_out(self):
        size_str = self.size
        if size_str:
            size_obj = Size(size_str)
            if size_obj.out_diameter is not None:
                return size_obj.out_diameter

        return self._get_param(
            "Main Secondary Diameter", unit=UnitTypeId.Inches, as_type="double")

    @property
    def height_in(self):
        size_str = self.size
        if size_str:
            size_obj = Size(size_str)
            if size_obj.in_height is not None:
                return size_obj.in_height

        return self._get_param(
            "Main Primary Depth", unit=UnitTypeId.Inches, as_type="double")

    @property
    def width_in(self):
        size_str = self.size
        if size_str:
            size_obj = Size(size_str)
            if size_obj.in_width is not None:
                return size_obj.in_width

        return self._get_param(
            "Main Primary Width", unit=UnitTypeId.Inches, as_type="double")

    @property
    def width_out(self):
        size_str = self.size
        if size_str:
            size_obj = Size(size_str)
            if size_obj.out_width is not None:
                return size_obj.out_width

        return self._get_param(
            "Main Secondary Width", unit=UnitTypeId.Inches, as_type="double")

    @property
    def height_out(self):
        size_str = self.size
        if size_str:
            size_obj = Size(size_str)
            if size_obj.out_height is not None:
                return size_obj.out_height

        return self._get_param(
            "Main Secondary Depth", unit=UnitTypeId.Inches, as_type="double")

    @property
    # Ex: TDF, S&D
    def connector_0_type(self):
        return self._get_param("NaviateDBS_Connector0_EndCondition")

    @property
    # Ex: TDF, S&D
    def connector_1_type(self):
        return self._get_param("NaviateDBS_Connector1_EndCondition")

    @property
    # Ex: TDF, S&D
    def connector_2_type(self):
        return self._get_param("NaviateDBS_Connector2_EndCondition")

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
            "NaviateDBS_D_Top Extension", unit=UnitTypeId.Inches, as_type="double")

    @property
    def extension_bottom(self):
        return self._get_param(
            "NaviateDBS_D_Bottom Extension", unit=UnitTypeId.Inches, as_type="double")

    @property
    def extension_right(self):
        return self._get_param(
            "NaviateDBS_D_Right Extension", unit=UnitTypeId.Inches, as_type="double")

    @property
    def extension_left(self):
        return self._get_param(
            "NaviateDBS_D_Left Extension", unit=UnitTypeId.Inches, as_type="double")

    @property
    def duty(self):
        return self._get_param("System Abbreviation")

    @property
    def family(self):
        fam = self._get_param("Family")
        if fam:
            return fam

        fam = self._get_param("NaviateDBS_Family")
        if fam:
            return fam

        return None

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
            "Weight", unit=UnitTypeId.PoundsMass, as_type="double")

    @property
    def service(self):
        return self._get_param("NaviateDBS_ServiceName")

    @property
    def inner_radius(self):
        return self._get_param("NaviateDBS_D_Inner Radius")

    @property
    def area(self):
        return self._get_param(
            "NaviateDBS_SheetMetalArea", unit=UnitTypeId.SquareFeet, as_type="double")

    @property
    def metal_area(self):
        return self._get_param(
            "NaviateDBS_SheetMetalArea", unit=UnitTypeId.SquareFeet, as_type="double")

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
    def joint_size(self):
        conn0 = (self.connector_0_type or "").strip()
        conn1 = (self.connector_1_type or "").strip()
        key = (self.family, conn0)

        if conn0 != conn1:
            return JointSize.INVALID

        threshold = CONNECTOR_THRESHOLDS.get(key)
        if threshold is None or self.length is None:
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
    def by_system_type(cls, doc, view, system_type_name):
        return [d for d in cls.all(doc, view) if d.element.LookupParameter(
            "System Type").AsString() == system_type_name]

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

        return run

    @staticmethod
    def create_duct_run(start_duct, doc, view):
        """Find all connected ducts/fittings that match both shape and size of the starting duct."""
        run = set()
        to_visit = [start_duct]
        visited = set()
        # Parse starting duct shape and size using Size.in_shape()
        start_size_obj = Size(str(start_duct.size))

        def shape_key_from_size(size_obj):
            """Create a comparable key from Size using inlet fields only."""
            shape = size_obj.in_shape()
            if shape == "round" and size_obj.in_diameter is not None:
                return ("round", round(size_obj.in_diameter, 2))
            if shape == "oval" and size_obj.in_oval_dia is not None and size_obj.in_oval_flat is not None:
                return ("oval", round(size_obj.in_oval_dia, 2), round(size_obj.in_oval_flat, 2))
            if shape == "rectangle" and size_obj.in_width is not None and size_obj.in_height is not None:
                return ("rect", round(size_obj.in_width, 2), round(size_obj.in_height, 2))
            return ("unknown", str(size_obj.in_size))

        start_shape = shape_key_from_size(start_size_obj)

        while to_visit:
            duct = to_visit.pop()
            if duct.id in visited:
                continue
            visited.add(duct.id)
            run.add(duct)
            for connector in duct.get_connectors():
                if not connector.IsConnected:
                    continue
                all_refs = list(connector.AllRefs)
                for ref in all_refs:
                    if ref and hasattr(ref, 'Owner'):
                        connected_elem = ref.Owner
                        # Only process fabrication parts
                        if not isinstance(connected_elem, FabricationPart):
                            continue
                        try:
                            connected_duct = RevitDuct(
                                doc, view, connected_elem)
                        except Exception:
                            continue
                        # Parse connected duct shape and size via Size.in_shape()
                        connected_size_obj = Size(str(connected_duct.size))
                        connected_shape = shape_key_from_size(
                            connected_size_obj)
                        # Match by normalized shape/size only (avoid string formatting mismatches)
                        if connected_shape == start_shape and connected_duct.id not in visited:
                            to_visit.append(connected_duct)
        return list(run)

    @staticmethod
    def create_duct_run_same_height(start_duct, doc, view, height_tolerance=0.01):
        """Find all connected ducts/fittings that match shape, size, AND z-axis height.

        The run will stop if:
        - Size changes
        - Shape changes
        - Z-axis height (centerline elevation) changes beyond tolerance

        Allows: X-Y offsetting, turns on X-Y plane
        Prevents: Vertical (Z-axis) offsetting beyond tolerance

        Args:
            start_duct: RevitDuct object to start from
            doc: Revit document
            view: Revit view
            height_tolerance: Tolerance in feet for z-axis differences (default 0.01 ft ≈ 0.12 inches)

        Returns:
            List of RevitDuct objects in the connected run at same height
        """
        run = set()
        to_visit = [start_duct]
        visited = set()
        start_size_obj = Size(str(start_duct.size))

        def shape_key_from_size(size_obj):
            """Create a comparable key from Size using inlet fields only."""
            shape = size_obj.in_shape()
            if shape == "round" and size_obj.in_diameter is not None:
                return ("round", round(size_obj.in_diameter, 2))
            if shape == "oval" and size_obj.in_oval_dia is not None and size_obj.in_oval_flat is not None:
                return ("oval", round(size_obj.in_oval_dia, 2), round(size_obj.in_oval_flat, 2))
            if shape == "rectangle" and size_obj.in_width is not None and size_obj.in_height is not None:
                return ("rect", round(size_obj.in_width, 2), round(size_obj.in_height, 2))
            return ("unknown", str(size_obj.in_size))

        def get_duct_z_coordinate(duct):
            """Extract Z-coordinate (elevation) from the duct's centerline.

            Prefer the location curve midpoint Z; fall back to inlet origin Z.
            """
            # Try centerline midpoint Z
            try:
                loc = duct.element.Location
                if hasattr(loc, 'Curve') and loc.Curve:
                    c = loc.Curve
                    p0 = c.GetEndPoint(0)
                    p1 = c.GetEndPoint(1)
                    return (p0.Z + p1.Z) / 2.0
            except Exception:
                pass
            # Fallback to inlet origin Z
            try:
                inlet_data, outlet_data = duct._inlet_outlet_from_revit_xyz()
                if inlet_data and 'origin' in inlet_data:
                    origin = inlet_data['origin']  # XYZ object
                    return origin.Z
            except Exception:
                pass
            return None

        start_shape = shape_key_from_size(start_size_obj)
        start_z = get_duct_z_coordinate(start_duct)

        while to_visit:
            duct = to_visit.pop()
            if duct.id in visited:
                continue
            visited.add(duct.id)
            run.add(duct)
            for connector in duct.get_connectors():
                if not connector.IsConnected:
                    continue
                all_refs = list(connector.AllRefs)
                for ref in all_refs:
                    if ref and hasattr(ref, 'Owner'):
                        connected_elem = ref.Owner
                        # Only process fabrication parts
                        if not isinstance(connected_elem, FabricationPart):
                            continue
                        try:
                            connected_duct = RevitDuct(
                                doc, view, connected_elem)
                        except Exception:
                            continue

                        # Check if already visited
                        if connected_duct.id in visited:
                            continue

                        # Parse connected duct shape and size
                        connected_size_obj = Size(str(connected_duct.size))
                        connected_shape = shape_key_from_size(
                            connected_size_obj)

                        # Check Z-axis height difference
                        connected_z = get_duct_z_coordinate(connected_duct)
                        z_difference = abs(
                            connected_z - start_z) if (connected_z is not None and start_z is not None) else None

                        # Match shape, size, AND z-axis height
                        if (connected_shape == start_shape and
                            z_difference is not None and
                                z_difference <= height_tolerance):
                            to_visit.append(connected_duct)
        return list(run)

    @staticmethod
    def parse_length_string(length_str):
        """
        Convert a Revit length string (e.g., "4' - 8\"", "2' - 4 23/32\"") to inches (float).
        Returns 0.0 if parsing fails.
        """
        if not length_str or not isinstance(length_str, str):
            return 0.0
        # Pattern: feet, inches, optional fraction
        pattern = r"(\d+)'\s*-\s*(\d+)?(?:\s+(\d+)/(\d+))?\s*\""
        cleaned = length_str.replace("’", "'").replace(
            "”", '"').replace("″", '"')
        match = re.match(pattern, cleaned)
        if not match:
            # Try to parse as a simple float
            try:
                return float(length_str)
            except Exception:
                return 0.0
        feet = int(match.group(1)) if match.group(1) else 0
        inches = int(match.group(2)) if match.group(2) else 0
        num = int(match.group(3)) if match.group(3) else 0
        denom = int(match.group(4)) if match.group(4) else 1
        fraction = float(num) / float(denom) if denom else 0
        total_inches = feet * 12 + inches + fraction
        return total_inches
