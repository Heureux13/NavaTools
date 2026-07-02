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

    def __init__(self, doc, view):
        self.doc = doc
        self.view = view
        self._tags = self._collect_tags()

    def _collect_tags(self):
        return (
            FilteredElementCollector(self.doc, self.view.Id)
            .OfCategory(self.TAG_CATEGORY)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    @staticmethod
    def _norm_tuple(values):
        return tuple((s or '').strip().lower() for s in (values or ()))

    def _iter_tags(self):
        results = []
        for tag in self._tags:
            rvt_tag = RevitElement(self.doc, self.view, tag)
            tag_family = rvt_tag.get_param(RVT_FAMILY, as_type='string')
            tag_family = (tag_family or '').strip().lower()
            results.append((tag, rvt_tag, tag_family))
        return results


    def get_all_annot(self, key_name=None, keywords=None):
        """Returns all tags in current view that do not match
        key_name or contain a word in keywords"""
        key_name = self._norm_tuple(key_name)
        keywords = self._norm_tuple(keywords)

        tags = []

        for tag, _rvt_tag, tag_family in self._iter_tags():

            is_exact = tag_family in key_name
            is_keyword = any(k in tag_family for k in keywords)

            if not (is_exact or is_keyword):
                tags.append(tag)

        return tags

    def get_tags_by_family(self, family=None):
        # Returns all tags that match family
        family = self._norm_tuple(family)

    tags = []

        for tag, _rvt_tag, tag_family in self._iter_tags():
            if tag_family in family:
                tags.append(tag)

        return tags

    def get_tags_by_family_and_type(self, family=None, tag_types=None):
        # Returns all tags that match family and type
        family = self._norm_tuple(family)
        tag_types = self._norm_tuple(tag_types)

        tags = []

        for tag, rvt_tag, tag_family in self._iter_tags():
            tag_type = rvt_tag.get_param(RVT_TYPE, as_type='string')
            tag_type = (tag_type or '').strip().lower()

            if tag_family in family and tag_type in tag_types:
                tags.append(tag)

        return tags
