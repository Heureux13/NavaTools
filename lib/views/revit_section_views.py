# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

from pyrevit import script, revit
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
)
from config.parameters_registry import (
    RVT_VIEW_NAME,
)
from revit.revit_element import RevitElement

# Variables
# =======================================================================


class SectionViews:
    TAG_CATEGORY = BuiltInCategory.OST_Viewers

    def __init__(self, doc=None, view=None):
        self.doc = doc
        self.view = view

    def _collect_views(self, doc, view):
        return (
            FilteredElementCollector(doc, view.Id)
            .OfCategory(self.TAG_CATEGORY)
            .WhereElementIsNotElementType()
            .ToElements()
        )

    def _norm_tuple(self, values):
        return tuple((s or '').strip().lower() for s in (values or ()))

    def get_sections_in_view(self, doc, view, key_name=None, keywords=None):
        key_name = self._norm_tuple(key_name)
        keywords = self._norm_tuple(keywords)

        api_views = self._collect_views(doc, view)

        views = []

        for elem in api_views:
            rvt_el = RevitElement(doc, view, elem)
            vname = rvt_el.get_param(RVT_VIEW_NAME, as_type='string')
            vname = (vname or '').strip().lower()

            is_exact = vname in key_name
            is_keyword = any(k in vname for k in keywords)

            if not (is_exact or is_keyword):
                views.append(elem)

        return views
