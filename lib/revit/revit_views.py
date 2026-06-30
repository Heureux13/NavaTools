# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

from pyrevit import revit
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    Viewport,
    XYZ,
    Transaction,
)
from config.parameters_registry import (
    RVT_VIEW_NAME,
)
from revit.revit_element import RevitElement

# Variables
# =======================================================================
default_coords = {
    'center': (10, 10),
}


class RevitViews:
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

    def get_views_in_view(self, doc, view, key_name=None, keywords=None):
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

    def get_viewport_info(self, doc, view):
        '''On sheet views, will take any selected views and return view id, view name, & center coordinates'''
        elements = revit.get_selection().elements
        viewports = [e for e in elements if isinstance(e, Viewport)]

        viewport_info = []

        for vp in viewports:
            center = vp.GetBoxCenter()
            outline = vp.GetBoxOutline()
            min_point = outline.MinimumPoint
            max_point = outline.MaximumPoint
            width = max_point.X - min_point.X
            height = max_point.Y - min_point.Y

            view_id = vp.ViewId
            view = doc.GetElement(view_id)

            rvt_el = RevitElement(doc, view, view)
            view_name = rvt_el.get_param(RVT_VIEW_NAME, as_type='string')

            viewport_info.append({
                'viewport_id': vp.Id.IntegerValue,
                'view_name': view_name,
                'center': (center.X, center.Y, center.Z),
                'width': width,
                'height': height,
                'max_point': max_point,
                'min_point': min_point,
            })

        return viewport_info

    def move_viewport_to_xyz(self, doc, default_coords, x=None, y=None):
        self.x = x if x is not None else default_coords['center'][0]
        self.y = y if y is not None else default_coords['center'][1]

        elements = revit.get_selection().elements
        viewports = [e for e in elements if isinstance(e, Viewport)]

        with Transaction(doc, 'Move Viewports') as txn:
            txn.Start()
            for vp in viewports:
                new_center = XYZ(self.x, self.y, 0)
                vp.SetBoxCenter(new_center)
            txn.Commit()
