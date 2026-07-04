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
    Reference,
    TagMode,
    TagOrientation,
    ElementId,
    XYZ,
)
from Autodesk.Revit.UI import UIDocument
from pyrevit import revit, forms, DB
from Autodesk.Revit.ApplicationServices import Application
from config.parameters_registry import RVT_FAMILY, RVT_TYPE, RVT_FAMILY_AND_TYPE
from config.tag_config import DEFAULT_TAG_SLOT_CANDIDATES, _load_user_candidate,
from revit.revit_element import RevitElement
import math

# Variables
# =======================================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument


# Functions
# =======================================================================


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
            rvt_tag = RevitElement(self.doc, self.view, tag)
            tag_family = rvt_tag.get_param(RVT_FAMILY, as_type="string")
            tag_type = rvt_tag.get_param(RVT_TYPE, as_type="string")
            tag_fam_n_type = rvt_tag.get_param(RVT_FAMILY_AND_TYPE, as_type="string")
            results.append((tag, tag_family, tag_type, tag_fam_n_type))
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

        raise LookupError("No label found with family '{}' and type '{}'".format(family_name, type_name))


    def get_tag_symbol_ids_from_slot_map(self, slot_map=None, slots=None):
        """Resolve config tag slots to loaded tag symbol type ids.

        Returns:
            dict[str, set[int]]: {slot_name: {tag_type_id_int, ...}}
        """
        candidates_by_slot = slot_map or DEFAULT_TAG_SLOT_CANDIDATES
        target_slots = slots if slots is not None else candidates_by_slot.keys()
        resolved = {}

        for slot in target_slots:
            type_ids = set()
            candidates = candidates_by_slot.get(slot, [])

            for family_name, type_name in candidates:
                try:
                    symbol = self._tag_symbol(family_name, type_name)
                    symbol_id = self._id_to_int(self._get_symbol_id(symbol))
                    if symbol_id is not None:
                        type_ids.add(symbol_id)
                except Exception:
                    continue

            resolved[slot] = type_ids

        return resolved


    # Class methods
    # =======================================================================================
    def create_tag(self,
                   element,
                   tag_symbol,
                   point_xyz):
        ref = self._ref(element)
        curve = element.Location.Curve

        return IndependentTag.Create(
            self.doc,
            tag_symbol.Id,
            self.view.Id,
            ref,
            False,
            TagOrientation.Horizontal,
            point_xyz,
        )

    def already_tagged(self,
                           element,
                           tag_symbol):
        """Returns true if an element as a matching tag symbol"""
        element = self._unwrap_element(element)

        if element is None or tag_symbol is None:
            return False

        tags = self.tags
        wanted_type_id = tag_symbol.GetTypeId()

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


    def get_tag_symbol_id_from_element(self,
                           element,):
       ele = self._unwrap_element(element)
       if ele is None:
           return set()

       symbol_ids = set()
       tags = self.tags

       for tag in tags:
           try:
               tagged_ids = self._get_ele_id_from_tag(tag)
               if ele.Id in tagged_ids:
                   symbol_id = self._id_to_int(tag.GetTypeId())
                   if symbol_id is not None:
                       symbol_ids.add(symbol_id)
           except Exception:
                continue

       return symbol_ids


    def midpoint_location(self,
                           element,
                           x_loc,
                           z_offset):
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

    def get_tag_symbol_id_from_family_and_type(self,
                                               family_name,
                                               type_name):
        tag_symbol = self._tag_symbol(family_name, type_name)
        return self._id_to_int(self._get_symbol_id(tag_symbol))


    def build_tag_symbol_id_map(self,
                                slot_map=None):
        if slot_map is None:
            slot_map = _load_user_candidate() or DEFAULT_TAG_SLOT_CANDIDATES

        slot_dict = {}

        for slot, candidate in slot_map.items():
            symbol_ids = []
            for family_name, type_name in candidate:
                try:
                    symbol_id = self.get_tag_symbol_id_from_family_and_type(family_name, type_name)
                    if symbol_id is not None:
                        symbol_ids.append(symbol_id)
                except Exception:
                    continue
            slot_dict[slot] = symbol_ids

        return

    def place_tag_rotated(self,
                          element,
                          tag_symbol,
                          position):

        loc = element.Location

        if not loc or not hasattr(loc, "Curve") or not loc.Curve:
            bbox = element.get_BoundingBox(self.view)
            if not bbox:
                return None
            center = (bbox.Min + bbox.Max) / 2.0
            tag = self.create_tag(element, tag_symbol, center)
            return tag

        curve = loc.Curve
        position_lower = self._clean(position)