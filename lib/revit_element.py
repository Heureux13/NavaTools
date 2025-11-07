# -*- coding: utf-8 -*-
"""
=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
=========================================================================
"""


from Autodesk.Revit.DB import *
from pyrevit import revit, forms, DB
from Autodesk.Revit.UI import UIDocument
from Autodesk.Revit.ApplicationServices import Application
import clr
from System.Collections.Generic import List

class RevitElement:
    def __init__(self, doc, view, element):
        self.doc = doc
        self.view = view
        self.element = element

    @property
    def id(self):
        return self.element.Id if self.element else None

    @property
    def category(self):
        return self.element.Category.Name if self.element and self.element.Category else None

    def get_param(self, param_name):
        param = self.element.LookupParameter(param_name)
        if param:
            if param.StorageType == DB.StorageType.String:
                return param.AsString()
            elif param.StorageType == DB.StorageType.Double:
                return param.AsDouble()
            elif param.StorageType == DB.StorageType.Integer:
                return param.AsInteger()
            elif param.StorageType == DB.StorageType.ElementId:
                return param.AsElementId()
        return None

    def set_param(self, param_name, value):
        param = self.element.LookupParameter(param_name)
        if not param:
            print("Parameter '{}' not found on element {}".format(param_name, self.id))
            return
        if param.IsReadOnly:
            print("Parameter '{}' is read-only on element {}".format(param_name, self.id))
            return

        if isinstance(value, str):
            param.Set(value)
        elif isinstance(value, int):
            param.Set(int(value))
        elif isinstance(value, float):
            param.Set(float(value))
        elif isinstance(value, DB.ElementId):
            param.Set(value)
        else:
            print("Unsupported value type for parameter '{}' on element {}".format(param_name, self.id))

    def select(self, uidoc, append=False):
        """Select this element in the Revit UI."""
        if not self.element:
            return

        if append:
            current = list(uidoc.Selection.GetElementIds())
            current.append(self.element.Id)
            id_list = List[ElementId](current)
        else:
            id_list = List[ElementId]([self.element.Id])

        uidoc.Selection.SetElementIds(id_list)

    @classmethod
    def select_many(cls, uidoc, elements):
        ids = List[ElementId]()
        for el in elements:
            if el is None:
                continue

            if hasattr(el, "element"):
                ids.Add(el.element.Id)

            elif isinstance(el, DB.Element):
                ids.Add(el.Id)

            elif isinstance(el, DB.ElementId):
                ids.Add(el)
                
        if ids.Count > 0:
            uidoc.Selection.SetElementIds(ids)

    # Print clickable links for each element and a 'Select All' link at the end.
    @staticmethod
    def print_select(output, elements, title="Title"):
        if not elements:
            output.print_md("No {} found.".format(title))
            return

        # Section title
        output.print_md("### {}".format(title))

        # Individual links
        for d in elements:
            output.print_md("- {}".format(output.linkify(d.id)))

        # Select All link
        all_ids = List[ElementId]()
        for d in elements:
            if d.id:
                all_ids.Add(d.id)

        output.print_md("**{}**".format(output.linkify(all_ids)))

        # Footer total
        output.print_md("**➡️{} of {} selected**".format(len(elements), title))