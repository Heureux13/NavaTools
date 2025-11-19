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
from revit_xyz import RevitXYZ
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
    def ogee_offset(self):
        return self._get_param("NaviateDBS_D_Offset", unit=UnitTypeId.Inches, as_type="double")

    @property
    def offset_width(self):
        return self._get_param("NaviateDBS_D_Offset-Width", unit=UnitTypeId.Inches, as_type="double")

    @property
    def offset_height(self):
        return self._get_param("NaviateDBS_D_Offset-Depth", unit=UnitTypeId.Inches, as_type="double")

    @property
    def width_in(self):
        return self._get_param("Main Primary Width", unit=UnitTypeId.Inches, as_type="double")

    @property
    def heigth_in(self):
        return self._get_param("Main Primary Depth", unit=UnitTypeId.Inches, as_type="double")

    @property
    def width_out(self):
        return self._get_param("Main Secondary Width", unit=UnitTypeId.Inches, as_type="double")

    @property
    def heigth_out(self):
        return self._get_param("Main Secondary Depth", unit=UnitTypeId.Inches, as_type="double")

    @property
    def connector_0(self):
        return self.get_connector(0)

    @property
    def connector_1(self):
        return self.get_connector(1)

    @property
    def connector_2(self):
        return self.get_connector(2)

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

    @property
    def offset_data(self):
        """Cache and return offset calculations for the duct."""
        if not hasattr(self, '_offset_data'):
            # Use identified inlet/outlet instead of raw connectors
            c_in, c_out = self.identify_inlet_outlet()

            if c_in and c_out:
                w_i = self.width_in
                h_i = self.heigth_in
                w_o = self.width_out or w_i
                h_o = self.heigth_out or h_i

                # Revit internal units (feet) -> inches
                p_in = (c_in.Origin.X * 12.0, c_in.Origin.Y *
                        12.0, c_in.Origin.Z * 12.0)
                p_out = (c_out.Origin.X * 12.0, c_out.Origin.Y *
                         12.0, c_out.Origin.Z * 12.0)

                # Get coordinate system from INLET
                try:
                    cs = c_in.CoordinateSystem
                    u_hat = (cs.BasisX.X, cs.BasisX.Y, cs.BasisX.Z)
                    v_hat = (cs.BasisY.X, cs.BasisY.Y, cs.BasisY.Z)
                except Exception:
                    u_hat = (1.0, 0.0, 0.0)
                    v_hat = (0.0, 1.0, 0.0)

                # Keep height axis pointing up in world space to stabilize top/bottom
                if v_hat[2] < 0.0:
                    u_hat = (-u_hat[0], -u_hat[1], -u_hat[2])
                    v_hat = (-v_hat[0], -v_hat[1], -v_hat[2])

                # Centerline offsets (inlet to outlet)
                delta = (p_out[0] - p_in[0], p_out[1] -
                         p_in[1], p_out[2] - p_in[2])
                width_offset = abs(RevitXYZ.dot(delta, u_hat))
                height_offset = abs(RevitXYZ.dot(delta, v_hat))

                # Edge offsets (inlet to outlet)
                edge_offsets = RevitXYZ.edge_diffs_whole_in(
                    p_in, w_i, h_i, p_out, w_o, h_o, u_hat, v_hat)

                self._offset_data = {
                    'centerline_width': width_offset,
                    'centerline_height': height_offset,
                    'edges': edge_offsets
                }
            else:
                self._offset_data = None

        return self._offset_data

    @property
    def centerline_width(self):
        """Centerline width offset in inches."""
        data = self.offset_data
        return data['centerline_width'] if data else None

    @property
    def centerline_height(self):
        """Centerline height offset in inches."""
        data = self.offset_data
        return data['centerline_height'] if data else None

    @property
    def offset_left(self):
        """Left edge offset in whole inches."""
        data = self.offset_data
        return data['edges']['whole_in']['left'] if data and data['edges'] else None

    @property
    def offset_right(self):
        """Right edge offset in whole inches."""
        data = self.offset_data
        return data['edges']['whole_in']['right'] if data and data['edges'] else None

    @property
    def offset_top(self):
        """Top edge offset in whole inches."""
        data = self.offset_data
        return data['edges']['whole_in']['top'] if data and data['edges'] else None

    @property
    def offset_bottom(self):
        """Bottom edge offset in whole inches."""
        data = self.offset_data
        return data['edges']['whole_in']['bottom'] if data and data['edges'] else None

    def connector_elevation(self, connector_index):
        """Get Z elevation of a connector in feet."""
        connector = self.get_connector(connector_index)
        return connector.Origin.Z if connector else None

    def higher_connector_index(self):
        """Return the index (0 or 1) of the higher connector, or None if can't determine."""
        c0 = self.get_connector(0)
        c1 = self.get_connector(1)

        if not c0 or not c1:
            return None

        z0 = c0.Origin.Z
        z1 = c1.Origin.Z

        if abs(z1 - z0) < 1e-6:  # essentially equal elevation
            return None

        return 1 if z1 > z0 else 0

    def is_connector_higher(self, connector_index, than_index):
        """Check if connector at connector_index is higher than connector at than_index."""
        c1 = self.get_connector(connector_index)
        c2 = self.get_connector(than_index)

        if not c1 or not c2:
            return None

        return c1.Origin.Z > c2.Origin.Z

    def top_edge_rise_in(self, tol_in=0.01):
        """Vertical rise (+) or drop (-) of top edge in inches between outlet and inlet."""
        c_in, c_out = self.identify_inlet_outlet()
        if not c_in or not c_out:
            return None

        h_in = self.heigth_in or 0.0
        h_out = (self.heigth_out or self.heigth_in or 0.0)

        try:
            cs_in = c_in.CoordinateSystem
            cs_out = c_out.CoordinateSystem
            v_in = cs_in.BasisY
            v_out = cs_out.BasisY
            vz_in = abs(v_in.Z)
            vz_out = abs(v_out.Z)

            top_in_z = c_in.Origin.Z + vz_in * (h_in / 2.0 / 12.0)
            top_out_z = c_out.Origin.Z + vz_out * (h_out / 2.0 / 12.0)
            rise_in = (top_out_z - top_in_z) * 12.0
        except:
            from Autodesk.Revit.DB import XYZ
            v_in = v_out = XYZ(0, 0, 1)

            top_in_z = c_in.Origin.Z + v_in.Z * (h_in / 2.0 / 12.0)
            top_out_z = c_out.Origin.Z + v_out.Z * (h_out / 2.0 / 12.0)
            rise_in = (top_out_z - top_in_z) * 12.0

        if abs(rise_in) < tol_in:
            return 0.0
        return round(rise_in, 2)

    def bottom_edge_rise_in(self, tol_in=0.01):
        """Vertical rise (+) or drop (-) of bottom edge in inches between outlet and inlet."""
        c_in, c_out = self.identify_inlet_outlet()
        if not c_in or not c_out:
            return None

        h_in = self.heigth_in or 0.0
        h_out = (self.heigth_out or self.heigth_in or 0.0)

        try:
            cs_in = c_in.CoordinateSystem
            cs_out = c_out.CoordinateSystem
            v_in = cs_in.BasisY
            v_out = cs_out.BasisY
            vz_in = abs(v_in.Z)
            vz_out = abs(v_out.Z)

            bottom_in_z = c_in.Origin.Z - vz_in * (h_in / 2.0 / 12.0)
            bottom_out_z = c_out.Origin.Z - vz_out * (h_out / 2.0 / 12.0)
            rise_in = (bottom_out_z - bottom_in_z) * 12.0
        except:
            from Autodesk.Revit.DB import XYZ
            v_in = v_out = XYZ(0, 0, 1)

            bottom_in_z = c_in.Origin.Z - v_in.Z * (h_in / 2.0 / 12.0)
            bottom_out_z = c_out.Origin.Z - v_out.Z * (h_out / 2.0 / 12.0)
            rise_in = (bottom_out_z - bottom_in_z) * 12.0

        if abs(rise_in) < tol_in:
            return 0.0
        return round(rise_in, 2)

    def identify_inlet_outlet(self):
        """Deterministic inlet/outlet by matching actual connector size to Primary (inlet) size."""
        try:
            conns = list(self.element.ConnectorManager.Connectors)
            if len(conns) < 2:
                return (None, None)
            c0, c1 = conns[0], conns[1]

            # Get parameter sizes (Primary = inlet, Secondary = outlet)
            w_primary = self.width_in
            h_primary = self.heigth_in

            # Try to get actual connector sizes
            try:
                # For rectangular connectors, check width/height
                w0 = c0.Width * 12.0  # feet to inches
                h0 = c0.Height * 12.0

                # If c0 size matches primary, it's the inlet
                if abs(w0 - w_primary) < 1.0 and abs(h0 - h_primary) < 1.0:
                    return (c0, c1)  # c0 = inlet, c1 = outlet
                else:
                    return (c1, c0)  # c1 = inlet, c0 = outlet
            except:
                # Fallback: assume c0 is inlet
                return (c0, c1)
        except Exception:
            return (None, None)
