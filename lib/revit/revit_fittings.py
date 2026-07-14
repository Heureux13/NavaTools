# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""
from config.tag_config import (
    DUCT_FAMILY_TAG_SLOTS,
    SLOT_EXT_TOP,
    SLOT_EXT_BOT,
    SLOT_EXT_LEFT,
    SLOT_EXT_RIGHT,
    RVT_ANGLE,
    RVT_LENGTH,
    NDBS_D_TOP_EXTENSION,
    NDBS_D_LEFT_EXTENSION,
    NDBS_D_RIGHT_EXTENSION,
    NDBS_D_BOTTOM_EXTENSION,
    NDBS_CONNECTOR0_END_CONDITION,
    NDBS_CONNECTOR1_END_CONDITION,
    NDBS_CONNECTOR2_END_CONDITION,
)
from config.duct_families import *
from revit.revit_element import RevitElement
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    IndependentTag,
    BuiltInCategory,
    FabricationPart,
)
from collections import defaultdict
import re


skip_values = {"skip"}

elbow_extensions_values = {
    # EXTENSION: (TDF value, S&D value)
    NDBS_D_TOP_EXTENSION: (6, 6),
    NDBS_D_LEFT_EXTENSION: (6, 6),
    NDBS_D_RIGHT_EXTENSION: (6, 6),
    NDBS_D_BOTTOM_EXTENSION: (6, 6),
}

angle_values = {
    RVT_ANGLE: (90, 45),
}


class RevitFittings:

    def __init__(self,
                 doc,
                 view,):
        self.doc = doc
        self.view = view
        self.collected_ducts = list(
            FilteredElementCollector(self.doc)
            .WhereElementIsNotElementType()
            .OfCategory(BuiltInCategory.OST_MEPFabrication)
            .ToElements()
        )
        self.collected_grds = list(
            FilteredElementCollector(self.doc)
            .WhereElementIsNotElementType()
            .OfCategory(BuiltInCategory.OST_DuctTerminal)
            .ToElements()
        )
        self.duct_map = self.create_duct_map()

        self.straights = {}
        self.elbows = {}
        self.reducers = {}
        self.ogees = {}
        self.square_rounds = {}
        self.taps = {}
        self.offsets = {}
        self.transitions = {}
        self.endcaps = {}
        self.dampers_volueme = {}
        self.canvas = {}
        self.access_panel = {}
        self.man_bars = {}

        for element_id, data in self.duct_map.items():
            family = data['family']

            if family == "elbows":
                self.elbows[element_id] = data


    @staticmethod
    def _clean(value):
        return (value or "").strip().lower()

    @staticmethod
    def _normalize_connector_type(connector_type):
        if not connector_type:
            return ''
        key = re.sub(r'\s+', ' ', str(connector_type).strip().lower())
        if 'tdf' in key:
            return 'tdf'
        if key in {'slip & drive', 'standing s&d', 'standing s and d', 's and d', 's&d'}:
            return 's&d'
        return key

    def _refresh_fabrications(self):
        self.collected_ducts = list(
            FilteredElementCollector(self.doc)
            .WhereElementIsNotElementType()
            .OfCategory(BuiltInCategory.OST_MEPFabrication)
            .ToElements()
        )

    def _family(self, fab_element):
        return self._clean(getattr(fab_element, "Name", ""))

    def _get_param(self,
                   fab_element,
                   param_name,
                   as_type=None,
                   unit=None):
        thing = RevitElement(self.doc, self.view, fab_element)
        return thing.get_param(param_name, as_type=as_type, unit=unit)

    def _skip_fab_element(self,
                          fab_element):
        value = self._get_param(fab_element, PYT_SKIP_TAG)
        if value is None:
            return None
        return self._clean(value) in {self._clean(v) for v in skip_values}


    def check_parameter(self,
                        fab_element,):
        ...

    def remove_tag_elements(self,
                            fab_element,):
        ...

    def tag_fab_element(self,
                        fab_element,
                        tag_element,):

    def create_family_map(self):
        groups = defaultdict(list)

        for fab_element in self.collected_ducts:
            family_name = self._family(fab_element)
            groups[family_name].append(fab_element)

        return groups

    def create_duct_map(self):
        element_map = {}

        for fab_element in self.collected_ducts:
            element_id = fab_element.Id

            element_map[element_id] = {
                "family": self._family(fab_element),
                "ext_bottom": self._get_param(fab_element, NDBS_D_BOTTOM_EXTENSION),
                "ext_top": self._get_param(fab_element, NDBS_D_TOP_EXTENSION),
                "ext_right": self._get_param(fab_element, NDBS_D_RIGHT_EXTENSION),
                "ext_left": self._get_param(fab_element, NDBS_D_LEFT_EXTENSION),
                "object": fab_element
            }

        return element_map



    def tag_elbows(self):
        ...
    def tag_straights(self):
        ...
    def tag_reducers(self):
        ...
    def tag_transitions(self):
        ...
    def tag_offsets(self):
        ...
    def tag_endcaps(self):
        ...
    def


    # TODO: build a map that separates them by family, that would speed up the parameter check.
