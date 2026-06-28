# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

from pyrevit import script, revit, forms, DB
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BoundingBoxXYZ,
    BuiltInCategory,
    ElementId,
    Category,
    View3D,
    ViewSection,
    ViewFamily,
    ViewFamilyType,
    XYZ,
    Transform,
    View,
)
from System.Collections.Generic import List
from config.parameters_registry import (
    RVT_FAMILY,
    RVT_FAMILY_AND_TYPE,
    RVT_TYPE,
    RVT_VIEW_NAME,
)
from revit.revit_element import RevitElement

# Variables
# =======================================================================
doc = revit.doc  # type: ignore[attr-defined]
uidoc = revit.uidoc  # type: ignore[attr-defined]
output = script.get_output()


class SectionViews:
    def __init__(self, doc=None, view=None, element=None):
        self.doc = doc
        self.view = view
        self.element = element

    @staticmethod
    def get_sections_on_view_old(doc, view):
        sections = []
        collector = (
            FilteredElementCollector(doc, view.Id)
            .WhereElementIsNotElementType()
        )
        for elem in collector.ToElements():
            try:
                cat = elem.Category
                if not cat or cat.Name != 'Views':
                    continue
            except Exception:
                continue

            # Confirm it is a section by checking name/family tokens.
            tokens = []
            try:
                name_text = getattr(elem, 'Name', None)
                if name_text:
                    tokens.append(str(name_text).lower())
            except Exception:
                pass
            for pname in (RVT_FAMILY, RVT_FAMILY_AND_TYPE, RVT_TYPE):
                try:
                    p = elem.LookupParameter(pname)
                    if p:
                        p_text = p.AsString() or p.AsValueString()
                        if p_text:
                            tokens.append(str(p_text).lower())
                except Exception:
                    pass
            if 'section' in ' '.join(tokens):
                sections.append(elem)

        return sections

    def get_sections_in_view(self, doc, view, key_name=None, keywords=None):
        key_name = tuple((s or '').strip().lower() for s in (key_name or ()))
        keywords = tuple((s or '').strip().lower() for s in (keywords or ()))

        api_views = (
            FilteredElementCollector(doc, view.Id)
            .OfCategory(BuiltInCategory.OST_Viewers)
            .WhereElementIsNotElementType()
            .ToElements()
        )

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
