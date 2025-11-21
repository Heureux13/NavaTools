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
from Autodesk.Revit.DB import (
    ElementId,
    FilteredElementCollector,
    BuiltInCategory,
    UnitUtils,
    FabricationPart,
    UnitTypeId
)
import re
import logging
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
    ("Straight", "TDC"): 56.25,
    ("Straight", "TDF"): 56.25,
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
    def connector_0_type(self):
        return self._get_param(
            "NaviateDBS_Connector0_EndCondition")

    @property
    def connector_1_type(self):
        return self._get_param(
            "NaviateDBS_Connector1_EndCondition")

    @property
    def connector_2_type(self):
        return self._get_param(
            "NaviateDBS_Connector2_EndCondition")

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
    def weight_metal(self):
        return self._get_param(
            "Weight",
            unit=UnitTypeId.PoundsMass,
            as_type="double")

    @property
    def service(self):
        return self._get_param("NaviateDBS_ServiceName")

    @property
    def inner_radius(self):
        return self._get_param("NaviateDBS_InnerRadius")

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
        """Returns a four option variance, one being an error. These sizes and connections can be changed easily across various fabs."""
        # Use the new properties that read from parameters
        conn0 = (self.connector_0_type or "").strip()
        conn1 = (self.connector_1_type or "").strip()

        if not conn0 or not conn1:
            return JointSize.INVALID

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

                # Keep height axis pointing up in world space to stabilize
                # top/bottom
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
        if data and data['edges']:
            return data['edges']['whole_in']['bottom']
        return None

    def connector_elevation(self, connector_index):
        """Get Z elevation of a connector in feet."""
        connector = self.get_connector(connector_index)
        return connector.Origin.Z if connector else None

    def higher_connector_index(self):
        """Return the index (0 or 1) of the higher connector.

        Returns None if can't determine."""
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
            d_primary = self.diameter_in

            # Try to get actual connector sizes
            try:
                # Convert values from feet to inches
                w0 = c0.Width * 12.0
                h0 = c0.Height * 12.0
                w1 = c1.Width * 12.0
                h1 = c1.Height * 12.0
                d0 = c0.Radius * 12.0 * 2
                d1 = c1.Radius * 12.0 * 2

                # Rectangular duct: compare by area (more robust than individual dimensions)
                if w_primary and h_primary:
                    area_primary = w_primary * h_primary
                    area_c0 = w0 * h0
                    area_c1 = w1 * h1

                    # Use area tolerance (5% of primary area)
                    area_tol = max(1.0, area_primary * 0.05)

                    # Check which connector's area is closer to primary area
                    diff_c0 = abs(area_c0 - area_primary)
                    diff_c1 = abs(area_c1 - area_primary)

                    if diff_c0 < area_tol and diff_c0 <= diff_c1:
                        return (c0, c1)  # c0 = inlet, c1 = outlet
                    elif diff_c1 < area_tol:
                        return (c1, c0)  # c1 = inlet, c0 = outlet

                # Round duct: compare by diameter
                if d_primary:
                    diff_c0 = abs(d0 - d_primary)
                    diff_c1 = abs(d1 - d_primary)
                    d_tol = max(0.5, d_primary * 0.05)

                    if diff_c0 < d_tol and diff_c0 <= diff_c1:
                        return (c0, c1)
                    elif diff_c1 < d_tol:
                        return (c1, c0)

            except BaseException:
                # Fallback: assume c0 is inlet
                return (c0, c1)

            # Fallback if no match found
            return (c0, c1)

        except Exception:
            return (None, None)

    def classify_offset(self):
        """Classify transition/reducer offset as CL/FOB/FOT/FOS or arrow/numeric offset.

        Returns dict with:
            - tag: str (CL, FOB, FOT, FOS, arrow like ↑2", or numeric like 3"→)
            - centerline_h: float (vertical centerline offset, inches)
            - centerline_w: float (horizontal centerline offset, inches)
            - top_edge: float (top edge rise, inches, signed)
            - bot_edge: float (bottom edge rise, inches, signed)
            - top_mag: int (top edge magnitude, whole inches)
            - bot_mag: int (bottom edge magnitude, whole inches)
        """
        c_in, c_out = self.identify_inlet_outlet()
        if not (c_in and c_out):
            return None

        p_in = c_in.Origin
        p_out = c_out.Origin

        # Horizontal centerline offset (plan distance)
        dx = p_out.X - p_in.X
        dy = p_out.Y - p_in.Y
        cen_w = (dx * dx + dy * dy) ** 0.5 * 12.0

        # Vertical centerline offset
        dz = p_out.Z - p_in.Z
        cen_h = abs(dz) * 12.0

        # Sizes
        h_in = self.heigth_in or self.diameter_in or 0.0
        h_out = self.heigth_out or self.diameter_out or h_in

        # World Z planes (feet)
        top_in_z = p_in.Z + 0.5 * (h_in / 12.0)
        top_out_z = p_out.Z + 0.5 * (h_out / 12.0)
        bot_in_z = p_in.Z - 0.5 * (h_in / 12.0)
        bot_out_z = p_out.Z - 0.5 * (h_out / 12.0)

        # Edge rises (inches, signed)
        top_e = (top_out_z - top_in_z) * 12.0
        bot_e = (bot_out_z - bot_in_z) * 12.0

        # Tolerance
        tol_in = 0.01

        top_aligned = abs(top_e) < tol_in
        bot_aligned = abs(bot_e) < tol_in
        cl_vert = top_aligned and bot_aligned

        # Whole-inch magnitudes
        off_t = int(round(abs(top_e)))
        off_b = int(round(abs(bot_e)))

        return {
            'centerline_w': cen_w,
            'centerline_h': cen_h,
            'top_edge': top_e,
            'bot_edge': bot_e,
            'top_mag': off_t,
            'bot_mag': off_b,
            'top_aligned': top_aligned,
            'bot_aligned': bot_aligned,
            'cl_vert': cl_vert
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
        offset_list = ["ogee", "offset", "radius offset", "mitered offset", "mitred offset"]
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

            if cl_vert or is_rotation:
                return "CL"
            elif bot_aligned:
                return "FOB"
            elif top_aligned:
                return "FOT"
            else:
                mag = int(round(abs(top_e)))
                if mag == 0:
                    return "CL"
                else:
                    return u'↑{}"'.format(mag) if top_e > 0 else u'↓{}"'.format(mag)

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
                    return u'{}"→'.format(int(round(y_off)))

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
