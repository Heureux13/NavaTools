# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ElementType,
    ElementTransformUtils,
    FamilySymbol,
    IndependentTag,
    Line,
    Reference,
    TagMode,
    TagOrientation,
    ElementId,
    XYZ,
)
from Autodesk.Revit.UI import UIDocument
from pyrevit import revit, forms, DB
from Autodesk.Revit.ApplicationServices import Application
from config.tag_config import DEFAULT_TAG_SLOT_CANDIDATES, _load_user_candidate
import math

# Variables
# =======================================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument


# Classes
# =======================================================================
class RevitTagging:
    def __init__(self, doc=None,
                 view=None):
        self.doc = doc or revit.doc
        self.view = view or revit.active_view
        self.fabduct_symbols = list(
            FilteredElementCollector(self.doc)
            .WhereElementIsElementType()
            .OfCategory(DB.BuiltInCategory.OST_FabricationDuctworkTags)
            .ToElements()
        )
        self.tags = (
            FilteredElementCollector(self.doc, self.view.Id)
            .OfClass(IndependentTag)
            .ToElements()
        )
        self._tag_data = self._iter_tag()

    # Static methods
    # =====================================================================
    @staticmethod
    def _clean(value):
        return (value or "").strip().lower()

    @staticmethod
    def _get_family_id(element):
        return element.Family.Id

    @staticmethod
    def _get_symbol_id(element):
        return element.GetTypeId()

    @staticmethod
    def _get_instance_id(element):
        return element.Id

    @staticmethod
    def _get_ele_id_from_tag(tag):
        if hasattr(tag, "GetTaggedLocalElementIds"):
            return list(tag.GetTaggedLocalElementIds() or [])
        if hasattr(tag, "TaggedLocalElementId"):
            tag_id = tag.TaggedLocalElementId
            return [tag_id] if tag_id else []
        return []

    @staticmethod
    def _id_to_int(eid):
        if eid is None:
            return None
        value = getattr(eid, "Value", None)
        if value is not None:
            return int(value)
        integer_value = getattr(eid, "IntegerValue", None)
        if integer_value is not None:
            return int(integer_value)
        return int(eid)

    # Helper methods
    # =============================================================================
    def _unwrap_element(self,
                        element):
        """Unwraps element from wrapper if needed"""
        return getattr(element, "element", element)

    def _refresh_tag(self):
        self.tags = list(
            FilteredElementCollector(self.doc, self.view.Id)
            .OfClass(IndependentTag)
            .ToElements()
        )
        return self.tags

    def _iter_tag(self):
        results = []

        for tag in self.fabduct_symbols:
            fam_name       = self._clean(getattr(getattr(tag, "Family", None), "Name", None))
            type_name      = self._clean(DB.Element.Name.GetValue(tag))
            tag_fam_n_type = fam_name + " " + type_name
            results.append((tag, fam_name, type_name, tag_fam_n_type))
        return results


    def _ref(self,element):
        """Handles either an Element or a Reference safely."""
        element = self._unwrap_element(element)
        return element if isinstance(element, Reference) else Reference(element)

    def _tag_symbol(self,
                    family_name,
                    type_name):

        fam_lower = self._clean(family_name)
        typ_lower = self._clean(type_name)

        for tag, tag_family, tag_type, tag_fam_n_type in self._tag_data:
            if self._clean(tag_family or '') == fam_lower and self._clean(tag_type or '') == typ_lower:
                return tag

        raise LookupError(
            "No label found with family '{}' and type '{}'"
            .format(family_name, type_name))


    # Class methods
    # =======================================================================================
    def midpoint_location(self,
                           element,
                           x_loc,
                           z_offset):
        """x must be a value between 0 and 1"""
        ele = self._unwrap_element(element)
        loc = ele.Location

        if hasattr(loc, "Curve") and loc.Curve:
            pt = loc.Curve.Evaluate(x_loc, True)
            return DB.XYZ(pt.X, pt.Y, pt.Z + z_offset)

        v = self.view
        bbox = ele.get_BoundingBox(v) if v else None

        if bbox:
            center = (bbox.Min + bbox.Max) / 2.0

            return DB.XYZ(center.X, center.Y, center.Z + z_offset)

        return None

    def get_tag_angle(self,
              tag_element,):

        ele   = self._unwrap_element(tag_element)
        if ele is None:
            return None

        loc = getattr(ele, "Location", None)
        curve = getattr(loc, "Curve", None)
        if curve is None:
            return None

        start = curve.GetEndPoint(0)
        end = curve.GetEndPoint(1)

        dx = end.X - start.X
        dy = end.Y - start.Y

        deg = math.degrees(math.atan2(dy, dx)) % 360.0

        # Fold 0..360 into 0..90
        d180 = deg % 180.0
        return min(d180, 180.0 - d180)

    def rotate_tag(self,
                   tag_element,
                   element):

        tag = self._unwrap_element(tag_element)
        ele = self._unwrap_element(element)

        if tag is None or ele is None:
            return None
        

        # compute safe tag angle (your function)
        angle_deg = self.get_tag_angle(ele)

        if angle_deg is None:
            return tag
        angle_rad = math.radians(angle_deg)

        # rotation axis: vertical line through tag head
        center = tag.TagHeadPosition
        axis = Line.CreateBound(center, center + XYZ(0, 0, 1))

        # rotate tag
        return ElementTransformUtils.RotateElement(
            self.doc,
            tag.Id,
            axis,
            angle_rad
        )

    def create_tag(self,
                   element,
                   tag_symbol,
                   orientation=None,
                   rotate=True,
                   x_loc=None,
                   z_loc=None):

        if x_loc is None:
            x_loc = 0.5
        if z_loc is None:
            z_loc = 0.0

        loc_point = self.midpoint_location(element, x_loc, z_loc)

        if loc_point is None:
            raise LookupError("No location found with x_loc, z_loc")

        ref = self._ref(element)
        ori = self._clean(orientation)

        if ori == "vertical":
            ori = TagOrientation.Vertical
        elif ori == "model":
            ori = TagOrientation.Model
        else:
            ori = TagOrientation.Horizontal

        tag = IndependentTag.Create(
            self.doc,
            tag_symbol.Id,
            self.view.Id,
            ref,
            False,
            ori,
            loc_point,
            )

        if rotate:
            self.rotate_tag(tag, element)

        return tag

    def already_tagged(self,
                           element,
                           tag_symbol):
        """Returns true if an element has a matching tag symbol"""
        element = self._unwrap_element(element)
        if element is None or tag_symbol is None:
            return False

        tags = self.tags
        wanted_type_id = tag_symbol.Id


        if wanted_type_id is None:
            return False

        for tag in tags:
            try:
                tagged_ids = self._get_ele_id_from_tag(tag)

                if element.Id in tagged_ids and tag.GetTypeId() == wanted_type_id:
                    return True

            except Exception:
                continue

        return False


    def get_tag_symbols_from_element(self,
                           element,):
       ele = self._unwrap_element(element)
       if ele is None:
           return []

       seen = {}

       for tag in self.tags:
           try:
               tagged_ids = self._get_ele_id_from_tag(tag)
               if ele.Id in tagged_ids:
                   symbol = self.doc.GetElement(tag.GetTypeId())
                   if symbol is not None:
                       seen[symbol.Id] = symbol
           except Exception:
                continue

       return list(seen.values())



    def build_tag_symbol_id_map(self, slot_map=None):
        if slot_map is None:
            slot_map = _load_user_candidate() or DEFAULT_TAG_SLOT_CANDIDATES

        # Precompute once from already-loaded tag data
        symbol_lookup = {}
        for tag, fam, typ, _ in self._tag_data:
            key = (self._clean(fam), self._clean(typ))
            sid = self._id_to_int(tag.Id)
            if sid is not None:
                symbol_lookup[key] = sid

        slot_dict = {}
        for slot, candidates in slot_map.items():
            ids = []
            for family_name, type_name in candidates:
                sid = symbol_lookup.get((self._clean(family_name), self._clean(type_name)))
                if sid is not None:
                    ids.append(sid)
            slot_dict[slot] = ids

        return slot_dict

