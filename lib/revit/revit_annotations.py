# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
)
from revit.revit_element import RevitElement
from config.parameters_registry import (
    RVT_FAMILY,
    RVT_TYPE,
)


class RevitAnnotations:
    TAG_CATEGORY = BuiltInCategory.OST_FabricationDuctworkTags

    def __init__(self, doc=None, view=None):
        self.doc = doc
        self.view = view

    def _collect_tags(self, doc, view):
        return (
            FilteredElementCollector(doc, view.Id)
            .OfCategory(self.TAG_CATEGORY)
            .WhereElementIsNotElementType()
            .ToElements()
        )

    def _norm_tuple(self, values):
        return tuple((s or '').strip().lower() for s in (values or ()))

    def get_all_annot(self, doc, view, key_name=None, keywords=None):
        key_name = self._norm_tuple(key_name)
        keywords = self._norm_tuple(keywords)

        api_duct_tags = self._collect_tags(doc, view)

        tags = []

        for tag in api_duct_tags:
            rvt_tag = RevitElement(doc, view, tag)
            tag_family = rvt_tag.get_param(RVT_FAMILY, as_type='string')
            tag_family = (tag_family or '').strip().lower()

            is_exact = tag_family in key_name
            is_keyword = any(k in tag_family for k in keywords)

            if not (is_exact or is_keyword):
                tags.append(tag)

        return tags

    def get_tags_by_family(self, doc, view, family=None):
        family = self._norm_tuple(family)

        api_duct_tags = self._collect_tags(doc, view)

        tags = []

        for tag in api_duct_tags:
            rvt_tag = RevitElement(doc, view, tag)

            tag_family = rvt_tag.get_param(RVT_FAMILY, as_type='string')
            tag_family = (tag_family or '').strip().lower()

            is_family = tag_family in family

            if is_family:
                tags.append(tag)

        return tags

    def get_tags_by_family_and_type(self, doc, view, family=None, tag_types=None):
        family = self._norm_tuple(family)
        tag_types = self._norm_tuple(tag_types)

        api_duct_tags = self._collect_tags(doc, view)

        tags = []

        for tag in api_duct_tags:
            rvt_tag = RevitElement(doc, view, tag)

            tag_family = rvt_tag.get_param(RVT_FAMILY, as_type='string')
            tag_family = (tag_family or '').strip().lower()

            tag_type = rvt_tag.get_param(RVT_TYPE, as_type='string')
            tag_type = (tag_type or '').strip().lower()

            is_family = tag_family in family
            is_type = tag_type in tag_types

            if (is_family and is_type):
                tags.append(tag)

        return tags
