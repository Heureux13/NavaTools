# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2026 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""
from config.tag_config import (
    DUCT_FAMILY_TAG_SLOTS,
    PYT_SKIP_TAG,
    PYT_OFFSET_VALUE,
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
from revit.revit_element import RevitElement, script
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
    NDBS_D_TOP_EXTENSION:       (6, 6),
    NDBS_D_LEFT_EXTENSION:      (6, 6),
    NDBS_D_RIGHT_EXTENSION:     (6, 6),
    NDBS_D_BOTTOM_EXTENSION:    (6, 6),
}

tdf_extension_value = 6
snd_extension_value = 6

angle_values = {
    RVT_ANGLE: (90, 45),
}


class RevitFittings:

    DUCT_PARAMETERS = {
        "ext_bottom":   NDBS_D_BOTTOM_EXTENSION,
        "ext_top":      NDBS_D_TOP_EXTENSION,
        "ext_right":    NDBS_D_RIGHT_EXTENSION,
        "ext_left":     NDBS_D_LEFT_EXTENSION,
        "angle":        RVT_ANGLE,
        "conn_0":       NDBS_CONNECTOR0_END_CONDITION,
        "conn_1":       NDBS_CONNECTOR1_END_CONDITION,
        "conn_2":       NDBS_CONNECTOR2_END_CONDITION,
        "offset":       PYT_OFFSET_VALUE,
    }

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

        family_map = self.create_family_map()

        self.straights = {}
        for fam in STRAIGHT_FAMILIES:
            self.straights.update(family_map.get(fam, {}))

        self.elbows = {}
        self.elbows_other = {}
        self.elbows_90 = {}
        self.elbows_45 = {}
        self.elbows_no_match = {}
        self.elbows_tdf = {}
        self.elbows_snd = {}
        self.elbows_mix_con = {}
        self.elbows_detag = {}

                for fam in ELBOW_FAMILIES:
            self.elbows.update(family_map.get(fam, {}))

        # Create dicts for elbows 45, 90, and anything else
        for element_id, data in self.elbows.items():
            angle = data.get("angle")

            try:
                angle = float(angle)
            except (TypeError, ValueError):
                self.elbows_no_match[element_id] = data
                continue

            if abs(angle - 90) < 0.5:
                self.elbows_90[element_id] = data
            elif abs(angle - 45) < 0.5:
                self.elbows_45[element_id] = data
            else:
                self.elbows_other[element_id] = data

        # Create dicts for elbows who's connectors are not 6"
        for element_id, data in self.elbows.items():
            conn_0 = self._clean(data.get("conn_0"))
            conn_1 = self._clean(data.get("conn_1"))

            if conn_0 == conn_1:
                if conn_0 in TDF_CONNECTOR_VALUES:
                    if conn_0 == tdf_extension_value:
                        self.elbows_tdf[element_id] = data
                if conn_0 in SND_CONNECTORS:
                    if conn_0 == snd_extension_value:
                        self.elbows_snd[element_id] = data

            else:
                self.elbows_mix_con[element_id] = data

        for element_id, data, in self.elbows_tdf.items():
            bottom_extension = data.get("bottom_extension")
            top_extension = data.get("top_extension")

            try:
                bottom_extension = float(bottom_extension)
                top_extension = float(top_extension)
            except (TypeError, ValueError):
                self.elbows_no_match[element_id] = data

                if bottom_extension == tdf_extension_value:
                    self.elbows_detag[element_id] = data
                else:
                    self.elbows_tag[element_id] = data





        self.taps = {}
        for fam in TAP_FAMILIES:
            self.taps.update(family_map.get(fam, {}))

        self.offsets = {}
        for fam in OFFSET_FAMILIES:
            self.offsets.update(family_map.get(fam, {}))

        self.endcaps = {}
        for fam in ENDCAP_FAMILIES:
            self.endcaps.update(family_map.get(fam, {}))

        self.canvas = {}
        for fam in CANVAS_FAMILIES:
            self.canvas.update(family_map.get(fam, {}))

        self.access_panel = {}
        for fam in ACCESS_DOOR_FAMILIES:
            self.access_panel.update(family_map.get(fam, {}))

        self.dampers = {}
        for fam in DAMPER_FAMILIES:
            self.dampers.update(family_map.get(fam, {}))


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

    def angle_matches(self, value, target, tol=0.5):
        try:
            return abs(float(value) - target) <= tol
        except (TypeError, ValueError):
            return False


    def check_parameter(self,
                        fab_element,):
        ...

    def remove_tag_elements(self,
                            fab_element,):
        ...

    def tag_fab_element(self,
                        fab_element,
                        tag_element,):
        ...

    def create_family_map(self):
        groups = defaultdict(dict)

        for element_id, data in self.duct_map.items():
            groups[data["family"]][element_id] = data

        return groups

    def create_duct_map(self):
        element_map = {}

        for fab_element in self.collected_ducts:
            element_data = {
                "family": self._family(fab_element),
                "object": fab_element,
            }

            for key, param_name in self.DUCT_PARAMETERS.items():
                element_data[key] = self._get_param(fab_element, param_name)

            element_map[fab_element.Id] = element_data

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


    # TODO: build a map that separates them by family, that would speed up the parameter check.
