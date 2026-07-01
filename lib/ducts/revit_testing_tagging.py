# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

from Autodesk.Revit.DB import (
    ElementId,
    StorageType,
    UnitUtils,
    FilteredElementCollector,
    BuiltInCategory,
)
from System.Collections.Generic import List
from revit.revit_element import RevitElement
from ducts.revit_duct import RevitDuct
from config.duct_families import duct_families
from config.parameters_registry import (
    RVT_FAMILY,
    RVT_TYPE,
)
import logging

# Global Variables
# ==========================================================
log = logging.getLogger("RevitElements")


class RevitTagging:
    CATEGORY_TAG = BuiltInCategory.OST_FabricationDuctworkTags
    CATEGORY_DUCT = BuiltInCategory.OST_FabricationDuctwork

    def __init__(self, doc, view, element, tag):
        self.doc = doc
        self.view = view
        self.element = element
        self.tag = tag

    @staticmethod
    def _norm_text(value):
        return (value or '').strip().lower()

    def _norm_tuple(self, value):
        return tuple((s or '').strip().lower() for s in (value or ()))

    def _collect_tags(self):
        return (
            FilteredElementCollector(self.doc, self.view.Id)
            .OfCategory(self.CATEGORY_TAG)
            .WhereElementIsNotElementType()
            .ToElements()
        )

    def _collect_ductwork(self):
        return (
            FilteredElementCollector(self.doc, self.view.Id)
            .OfCategory(self.CATEGORY_DUCT)
            .WhereElementIsNotElementType()
            .ToElements()
        )

    def get_all_tags(self, key_name=None, keywords=None):
        key_name = self._norm_tuple(key_name)
        keywords = self._norm_tuple(keywords)

        api_duct_tags = self._collect_tags()

        tags = []

        for tag in api_duct_tags:
            rvt_tag = RevitElement(self.doc, self.view, tag)
            tag_family = rvt_tag.get_param(RVT_FAMILY, as_type='string')
            tag_family = (tag_family or '').strip().lower()

            is_exact = tag_family in key_name
            is_keyword = any(k in tag_family for k in keywords)

            if not (is_exact or is_keyword):
                tags.append(tag)

        return tags

    def get_tags_by_family(self, family=None):
        family = self._norm_tuple(family)

        api_duct_tags = self._collect_tags()

        tags = []

        for tag in api_duct_tags:
            rvt_tag = RevitElement(self.doc, self.view, tag)

            tag_family = rvt_tag.get_param(RVT_FAMILY, as_type='string')
            tag_family = (tag_family or '').strip().lower()

            is_family = tag_family in family

            if is_family:
                tags.append(tag)

        return tags

    def get_tags_by_family_and_type(self, family=None, tag_types=None):
        family = self._norm_tuple(family)
        tag_types = self._norm_tuple(tag_types)

        api_duct_tags = self._collect_tags()

        tags = []

        for tag in api_duct_tags:
            rvt_tag = RevitElement(self.doc, self.view, tag)

            tag_family = rvt_tag.get_param(RVT_FAMILY, as_type='string')
            tag_family = (tag_family or '').strip().lower()

            tag_type = rvt_tag.get_param(RVT_TYPE, as_type='string')
            tag_type = (tag_type or '').strip().lower()

            is_family = tag_family in family
            is_type = tag_type in tag_types

            if (is_family and is_type):
                tags.append(tag)

        return tags

    def get_duct_family(self):
        ducts = self._collect_ductwork()
        clean_dict = {(k or '').strip().lower(): v for k,
                      v in duct_families.items()}

        family_to_ids = {}

        for d in ducts:
            d_duct = RevitDuct(self.doc, self.view, d)
            fam = (d_duct.family or '').strip().lower()

            if fam not in clean_dict:
                continue
            else:
                family_to_ids.setdefault(fam, []).append(d.Id)

        return family_to_ids
