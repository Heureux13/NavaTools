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

    def _get_param(
            self, name, unit=None, as_type="string", required=False):
        # helper gettin parameters from revit
        p = self.element.LookupParameter(name)
        if not p:
            if required:
                raise KeyError(
                    "Missing parameter '{}' on element {}".format(
                        name, self.element.Id))
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
            # convert any unexpected Revit exception into None to keep callers
            # deterministic
            return None

    @property
    def size(self):
        return self._get_param("Size")

    @property
    def centerline_length(self):
        # Return centerline length as a numeric value in inches
        return self._get_param(
            "NaviateDBS_CenterlineLength",
            unit=UnitTypeId.Inches,
            as_type="double"
        )

    @property
    def length(self):
        return self._get_param(
            "Length",
            unit=UnitTypeId.Inches,
            as_type="double")

    @property
    def ogee_offset(self):
        return self._get_param(
            "NaviateDBS_D_Offset",
            unit=UnitTypeId.Inches,
            as_type="double")

    @property
    def reducer_offset(self):
        return self._get_param(
            "NaviateDBS_D_Y-Offset",
            unit=UnitTypeId.Inches,
            as_type="double")

    @property
    def offset_width(self):
        return self._get_param(
            "NaviateDBS_D_Offset-Width",
            unit=UnitTypeId.Inches,
            as_type="double")

    @property
    def offset_height(self):
        return self._get_param(
            "NaviateDBS_D_Offset-Depth",
            unit=UnitTypeId.Inches,
            as_type="double")

    @property
    def diameter_in(self):
        return self._get_param(
            "Main Primary Diameter",
            unit=UnitTypeId.Inches,
            as_type="double")

    @property
    def diameter_out(self):
        return self._get_param(
            "Main Secondary Diameter",
            unit=UnitTypeId.Inches,
            as_type="double")

    @property
    def width_in(self):
        return self._get_param(
            "Main Primary Width",
            unit=UnitTypeId.Inches,
            as_type="double")

    @property
    def heigth_in(self):
        return self._get_param(
            "Main Primary Depth",
            unit=UnitTypeId.Inches,
            as_type="double")

    @property
    def width_out(self):
        return self._get_param(
            "Main Secondary Width",
            unit=UnitTypeId.Inches,
            as_type="double")

    @property
    def heigth_out(self):
        return self._get_param(
            "Main Secondary Depth",
            unit=UnitTypeId.Inches,
            as_type="double")

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
            "NaviateDBS_D_Top Extension",
            unit=UnitTypeId.Inches,
            as_type="double")

    @property
    def extension_bottom(self):
        return self._get_param(
            "NaviateDBS_D_Bottom Extension",
            unit=UnitTypeId.Inches,
            as_type="double")

    @property
    def extension_right(self):
        return self._get_param(
            "NaviateDBS_D_Right Extension",
            unit=UnitTypeId.Inches,
            as_type="double")

    @property
    def extension_left(self):
        return self._get_param(
            "NaviateDBS_D_Left Extension",
            unit=UnitTypeId.Inches,
            as_type="double")

    @property
    def duty(self):
        return self._get_param("System Abbreviation")

    @property
    def family(self):
        fam = self._get_param("NaviateDBS_Family")
        if fam:
            return fam
        fam = self._get_param("Family")
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
            "Weight",
            unit=UnitTypeId.PoundsMass,
            as_type="double")

    @property
    def service(self):
        return self._get_param("NaviateDBS_ServiceName")

    @property
    def inner_radius(self):
        return self._get_param("NaviateDBS_D_Inner Radius")

    @property
    def area(self):
        return self._get_param(
            "NaviateDBS_SheetMetalArea",
            unit=UnitTypeId.SquareFeet,
            as_type="double")

    @property
    def metal_area(self):
        return self._get_param(
            "NaviateDBS_SheetMetalArea",
            unit=UnitTypeId.SquareFeet,
            as_type="double")

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

    @property
    def offset_data(self):
        """Cache and return offset calculations for the duct."""
        if not hasattr(self, '_offset_data'):
            # Use identified inlet/outlet instead of raw connectors
            c_in, c_out = self.identify_inlet_outlet()

            if c_in and c_out:
                # Detect round connectors (prefer explicit connector properties)
                def has_radius(conn):
                    try:
                        return hasattr(conn, 'Radius') and conn.Radius and conn.Radius > 1e-6
                    except Exception:
                        return False

                is_round_in = has_radius(c_in)
                is_round_out = has_radius(c_out)
                is_round = bool(is_round_in and is_round_out)

                # Get dimensions based on shape
                if is_round:
                    # For round: use diameter from connector or parameters
                    w_i = None
                    w_o = None
                    try:
                        r_in = c_in.Radius
                        if r_in and r_in > 1e-6:
                            w_i = r_in * 24.0
                    except Exception:
                        pass
                    if not w_i:
                        w_i = self.diameter_in

                    try:
                        r_out = c_out.Radius
                        if r_out and r_out > 1e-6:
                            w_o = r_out * 24.0
                    except Exception:
                        pass
                    if not w_o:
                        w_o = self.diameter_out
                    if not w_o:
                        w_o = w_i

                    h_i = w_i
                    h_o = w_o
                else:
                    # For rectangular: use width/height parameters
                    w_i = self.width_in
                    h_i = self.heigth_in
                    w_o = self.width_out or w_i
                    h_o = self.heigth_out or h_i

                # Validate we have dimensions
                if not w_i or not h_i:
                    self._offset_data = None
                    return self._offset_data

                # Revit internal units (feet) -> inches
                p_in = (c_in.Origin.X * 12.0, c_in.Origin.Y *
                        12.0, c_in.Origin.Z * 12.0)
                p_out = (c_out.Origin.X * 12.0, c_out.Origin.Y *
                         12.0, c_out.Origin.Z * 12.0)

                # Get coordinate system from INLET (cache to avoid repeated access)
                try:
                    cs = c_in.CoordinateSystem
                    bx = cs.BasisX
                    by = cs.BasisY
                    u_hat = (bx.X, bx.Y, bx.Z)
                    v_hat = (by.X, by.Y, by.Z)
                except Exception:
                    u_hat = (1.0, 0.0, 0.0)
                    v_hat = (0.0, 1.0, 0.0)

                # Keep height axis pointing up in world space to stabilize
                # top/bottom
                if v_hat[2] < 0.0:
                    u_hat = (-u_hat[0], -u_hat[1], -u_hat[2])
                    v_hat = (-v_hat[0], -v_hat[1], -v_hat[2])

                # Centerline offsets (inlet to outlet)
                delta = (
                    p_out[0] - p_in[0], p_out[1] - p_in[1], p_out[2] - p_in[2])
                width_offset = abs(RevitXYZ.dot(delta, u_hat))
                height_offset = abs(RevitXYZ.dot(delta, v_hat))

                # Edge offsets (inlet to outlet)
                if not is_round:
                    edge_offsets = RevitXYZ.edge_diffs_whole_in(
                        p_in, w_i, h_i, p_out, w_o, h_o, u_hat, v_hat)
                else:
                    # Round parts do not have meaningful rectangular edges.
                    # Provide None for edge offsets and include diameters for context.
                    try:
                        d_in = c_in.Radius * 24.0
                    except Exception:
                        d_in = self.diameter_in
                    try:
                        d_out = c_out.Radius * 24.0
                    except Exception:
                        d_out = self.diameter_out or d_in
                    edge_offsets = {
                        'whole_in': {
                            'left': None,
                            'right': None,
                            'top': None,
                            'bottom': None,
                        },
                        'round': True,
                        'diam_in': d_in,
                        'diam_out': d_out,
                    }

                self._offset_data = {
                    'centerline_width': width_offset,
                    'centerline_height': height_offset,
                    'edges': edge_offsets
                }
            else:
                # No valid connectors: cannot compute geometry-based offsets
                self._offset_data = None

        return self._offset_data

    @property
    def is_round(self):
        """True if both connectors are round (edge offsets not meaningful)."""
        data = getattr(self, '_offset_data', None)
        if not data:
            # Force calculation once if missing
            data = self.offset_data
        return bool(data and data.get('edges') and data['edges'].get('round'))

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
        if data and data['edges']:
            return data['edges']['whole_in']['bottom']
        return None

    def identify_inlet_outlet(self):
        """Deterministically pick inlet (larger connector) and outlet (smaller)."""
        try:
            conns = list(self.element.ConnectorManager.Connectors)
            if len(conns) < 2:
                return (None, None)
            c0, c1 = conns[0], conns[1]

            # Try rectangular sizes (inches)
            def rect_wh(conn):
                try:
                    return conn.Width * 12.0, conn.Height * 12.0
                except Exception:
                    return None, None

            w0, h0 = rect_wh(c0)
            w1, h1 = rect_wh(c1)

            # Try round diameters (inches)
            def diameter(conn):
                try:
                    return conn.Radius * 24.0  # 2 * radius * 12
                except Exception:
                    return None
            d0 = diameter(c0)
            d1 = diameter(c1)

            # Rectangular case first
            if w0 and h0 and w1 and h1:
                a0 = w0 * h0
                a1 = w1 * h1
                if abs(a0 - a1) > 1e-6:
                    return (c0, c1) if a0 >= a1 else (c1, c0)
                # Tie: fall back to element id for stability
                id0 = get_element_id_value(c0.Owner.Id)
                id1 = get_element_id_value(c1.Owner.Id)
                return (c0, c1) if id0 <= id1 else (c1, c0)

            # Round case
            if d0 and d1:
                if abs(d0 - d1) > 1e-6:
                    return (c0, c1) if d0 >= d1 else (c1, c0)
                id0 = get_element_id_value(c0.Owner.Id)
                id1 = get_element_id_value(c1.Owner.Id)
                return (c0, c1) if id0 <= id1 else (c1, c0)

            # Mixed shape case: one rectangular, one round
            # Compute areas and compare
            a0 = None
            a1 = None
            if w0 and h0:
                a0 = w0 * h0
            elif d0:
                a0 = 3.14159 * (d0 / 2.0) ** 2

            if w1 and h1:
                a1 = w1 * h1
            elif d1:
                a1 = 3.14159 * (d1 / 2.0) ** 2

            if a0 is not None and a1 is not None and abs(a0 - a1) > 1e-6:
                return (c0, c1) if a0 >= a1 else (c1, c0)

            # Last resort: use connector index for deterministic ordering
            # (both connectors belong to same element, so Owner.Id won't help)
            return (c0, c1)  # c0 is always inlet for consistency

        except Exception:
            return (None, None)

    def classify_offset(self):
        """Classify transition/reducer offset as CL/FOB/FOT/FOS or arrow/numeric offset."""
        c_in, c_out = self.identify_inlet_outlet()
        if not (c_in and c_out):
            return None

        p_in = c_in.Origin
        p_out = c_out.Origin

        # Vector from inlet to outlet
        delta = p_out - p_in

        # Get width direction (BasisX) from inlet connector
        # BasisX points along the width of rectangular duct (left-right direction)
        try:
            width_dir = c_in.CoordinateSystem.BasisX
            # Project delta onto width direction to get signed horizontal offset
            # Positive = offset in +BasisX direction (right), Negative = offset in -BasisX direction (left)
            offset_perp_signed = (delta.X * width_dir.X +
                                  delta.Y * width_dir.Y +
                                  delta.Z * width_dir.Z) * 12.0
            offset_perp = abs(offset_perp_signed)
        except Exception:
            # Fallback: no horizontal offset
            offset_perp_signed = 0.0
            offset_perp = 0.0
            offset_perp = 0.0

        # Horizontal centerline offset (plan distance - for reference)
        cen_w = math.hypot(delta.X, delta.Y) * 12

        # Vertical centerline offset (signed and magnitude)
        cen_h_signed = delta.Z * 12.0
        cen_h = abs(cen_h_signed)

        # Detect connector shapes
        def is_round(conn):
            try:
                return hasattr(conn, 'Radius') and conn.Radius and conn.Radius > 1e-6
            except Exception:
                return False

        round_in = is_round(c_in)
        round_out = is_round(c_out)

        # Handle mixed transitions (square to round)
        if round_in != round_out:
            # Mixed transition: use centerline offsets and perpendicular offset
            return {
                'centerline_w': cen_w,
                'centerline_h': cen_h,
                'centerline_h_signed': delta.Z * 12.0,
                'offset_perp': offset_perp,
                'offset_perp_signed': offset_perp_signed,
                'top_edge': None,
                'bot_edge': None,
                'left_edge': None,
                'right_edge': None,
                'top_mag': None,
                'bot_mag': None,
                'left_mag': None,
                'right_mag': None,
                'top_aligned': False,
                'bot_aligned': False,
                'left_aligned': False,
                'right_aligned': False,
                'cl_vert': True,
                'is_mixed': True
            }

        # Sizes (both width and height)
        w_in = c_in.Width * \
            12.0 if hasattr(c_in, 'Width') and c_in.Width else 0.0
        w_out = c_out.Width * \
            12.0 if hasattr(c_out, 'Width') and c_out.Width else w_in

        h_in = (c_in.Height * 12.0 if hasattr(c_in, 'Height') and c_in.Height
                else (c_in.Radius * 24.0 if hasattr(c_in, 'Radius') and c_in.Radius else 0.0))
        h_out = (c_out.Height * 12.0 if hasattr(c_out, 'Height') and c_out.Height
                 else (c_out.Radius * 24.0 if hasattr(c_out, 'Radius') and c_out.Radius else h_in))

        # World Z planes (feet) - using actual connector positions
        top_in_z = p_in.Z + 0.5 * (h_in / 12.0)
        top_out_z = p_out.Z + 0.5 * (h_out / 12.0)
        bot_in_z = p_in.Z - 0.5 * (h_in / 12.0)
        bot_out_z = p_out.Z - 0.5 * (h_out / 12.0)

        # Edge rises (inches, signed)
        top_e = (top_out_z - top_in_z) * 12.0
        bot_e = (bot_out_z - bot_in_z) * 12.0

        # Left and right edge offsets (if rectangular)
        left_in_z = p_in.Z - 0.5 * (w_in / 12.0)
        left_out_z = p_out.Z - 0.5 * (w_out / 12.0)
        right_in_z = p_in.Z + 0.5 * (w_in / 12.0)
        right_out_z = p_out.Z + 0.5 * (w_out / 12.0)

        left_e = (left_out_z - left_in_z) * 12.0
        right_e = (right_out_z - right_in_z) * 12.0

        # Tolerance
        tol_in = 0.01

        top_aligned = abs(top_e) < tol_in
        bot_aligned = abs(bot_e) < tol_in
        left_aligned = abs(left_e) < tol_in
        right_aligned = abs(right_e) < tol_in
        cl_vert = top_aligned and bot_aligned

        # Whole-inch magnitudes
        off_t = int(round(abs(top_e)))
        off_b = int(round(abs(bot_e)))
        off_l = int(round(abs(left_e)))
        off_r = int(round(abs(right_e)))

        return {
            'centerline_w': cen_w,
            'centerline_h': cen_h,
            'centerline_h_signed': cen_h_signed,
            'offset_perp': offset_perp,
            'offset_perp_signed': offset_perp_signed,
            'top_edge': top_e,
            'bot_edge': bot_e,
            'left_edge': left_e,
            'right_edge': right_e,
            'top_mag': off_t,
            'bot_mag': off_b,
            'left_mag': off_l,
            'right_mag': off_r,
            'top_aligned': top_aligned,
            'bot_aligned': bot_aligned,
            'left_aligned': left_aligned,
            'right_aligned': right_aligned,
            'cl_vert': cl_vert,
            'is_mixed': False,
            'w_in': w_in,
            'w_out': w_out,
            'h_in': h_in,
            'h_out': h_out
        }

    def get_offset_value(self):
        """Calculate offset classification tag for transitions/reducers/offsets.

        Returns:
            str: Tag like "CL", "FOB", "FOT", "FOS", "↑2"", "3"→", or None if not applicable.
        """
        family = (self.family or "").lower().strip()

        # Family lists
        reducer_square = ["transition"]
        reducer_round = ["reducer"]
        offset_list = ["ogee", "offset", "radius offset",
                       "mitered offset", "mitred offset"]
        family_list = reducer_square + reducer_round + offset_list

        if family not in family_list:
            return None

        # Get offset data
        offset_data = self.classify_offset()
        if not offset_data:
            return None

        cen_w = offset_data['centerline_w']
        cen_h = offset_data['centerline_h']
        top_e = offset_data['top_edge']
        bot_e = offset_data['bot_edge']
        top_aligned = offset_data['top_aligned']
        bot_aligned = offset_data['bot_aligned']
        cl_vert = offset_data['cl_vert']

        # Rectangular reducers/transitions
        if family in reducer_square:
            is_rotation = (cen_h < 0.5) and abs(abs(top_e) - abs(bot_e)) < 0.5

            # Get left/right edge data
            left_e = offset_data.get('left_edge', 0)
            right_e = offset_data.get('right_edge', 0)
            left_aligned = offset_data.get('left_aligned', False)
            right_aligned = offset_data.get('right_aligned', False)

            if cl_vert or is_rotation:
                return "CL"

            # Build combined tag for aligned edges
            aligned_edges = []
            if bot_aligned:
                aligned_edges.append("FOB")
            if top_aligned:
                aligned_edges.append("FOT")
            if left_aligned:
                aligned_edges.append("FOL")
            if right_aligned:
                aligned_edges.append("FOR")

            if aligned_edges:
                return "/".join(aligned_edges)

            # No edges aligned - show offsets with arrows
            # Check if both vertical AND horizontal offsets exist
            has_vert = abs(top_e) >= 0.5
            has_horiz = abs(left_e) >= 0.5 or abs(right_e) >= 0.5

            if has_vert and has_horiz:
                # Both directions - show both with space
                vert_mag = int(round(abs(top_e)))
                horiz_mag = int(round(abs(left_e)))
                vert_str = u'↑{}"TU'.format(
                    vert_mag) if top_e > 0 else u'↓{}"TD'.format(vert_mag)
                horiz_str = u'←{}"'.format(
                    horiz_mag) if left_e < 0 else u'→{}"'.format(horiz_mag)
                return u'{} {}'.format(vert_str, horiz_str)
            elif has_vert:
                # Only vertical
                mag = int(round(abs(top_e)))
                return u'↑{}"TU'.format(mag) if top_e > 0 else u'↓{}"TD'.format(mag)
            elif has_horiz:
                # Only horizontal
                mag = int(round(abs(left_e)))
                return u'←{}"'.format(mag) if left_e < 0 else u'→{}"'.format(mag)
            else:
                return "CL"

        # Round reducers
        elif family in reducer_round:
            y_off = self.reducer_offset
            d_in = self.diameter_in
            d_out = self.diameter_out

            if (y_off is not None) and (d_in is not None) and (d_out is not None):
                expected_cl = (d_in - d_out) / 2.0

                if abs(y_off - expected_cl) < 0.01:
                    return "CL"
                elif abs(d_out + y_off - d_in) < 0.01 or abs(y_off) < 0.1:
                    return "FOS"
                else:
                    return u'{}"→'.format(abs(int(round(y_off))))

        # Horizontal offsets
        elif family in offset_list:
            oge_o = self.ogee_offset
            offset = oge_o or cen_w or 0
            return u'{}"→'.format(int(round(offset)))

        return None

    def get_connected_elements(self, connector_index=0):
        """Gets all elements connected to the selected element"""
        connector = self.get_connector(connector_index)
        connected_elements = []

        if connector and connector.IsConnected:
            for ref_conn in connector.AllRefs:
                if ref_conn.Owner.Id != self.element.Id:
                    connected_elements.append(ref_conn.Owner)
        return connected_elements

    @staticmethod
    def trace_run(start_duct, seen_ids, allowed_duct, doc, view):
        """Recursively follow connections to build a complete run"""
        run = []  # Store ducts in this run
        stack = [start_duct]  # Ducts to process

        while stack:
            current = stack.pop()
            if current.id in seen_ids:
                continue

            run.append(current)
            seen_ids.add(current.id)

            # Follow all connections
            for connector_index in [0, 1, 2]:
                connected = current.get_connected_elements(connector_index)
                for elem in connected:
                    duct = RevitDuct(doc, view, elem)
                    if duct.family and duct.family.strip().lower() in allowed_duct:
                        stack.append(duct)

        return run

    @staticmethod
    def create_duct_run(start_duct, doc, view):
        """Find all connected ducts/fittings that match both shape and size of the starting duct."""
        run = set()
        to_visit = [start_duct]
        visited = set()
        # Parse starting duct shape and size
        start_size_obj = Size(str(start_duct.size))

        def get_shape_key(size_obj):
            # Returns a tuple representing the shape (round, oval, rectangular) and main dimensions
            if size_obj.in_diameter:
                return ("round", round(size_obj.in_diameter, 2))
            elif size_obj.in_oval_dia and size_obj.in_oval_flat is not None:
                return ("oval", round(size_obj.in_oval_dia, 2), round(size_obj.in_oval_flat, 2))
            elif size_obj.in_width and size_obj.in_height:
                return ("rect", round(size_obj.in_width, 2), round(size_obj.in_height, 2))
            else:
                return ("unknown", str(size_obj.in_size))

        start_shape = get_shape_key(start_size_obj)
        start_size_str = str(start_size_obj.in_size)

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
                        try:
                            connected_duct = RevitDuct(
                                doc, view, connected_elem)
                        except Exception:
                            continue
                        # Parse connected duct shape and size
                        connected_size_obj = Size(
                            str(connected_duct.size))
                        connected_shape = get_shape_key(connected_size_obj)
                        connected_size_str = str(connected_size_obj.in_size)
                        # Match both shape and size
                        if connected_shape == start_shape and connected_size_str == start_size_str and connected_duct.id not in visited:
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
