# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

# Imports
# =========================================================================
from Autodesk.Revit.DB import ElementId, StorageType, UnitUtils
from System.Collections.Generic import List
import logging

# Global Variables
# =========================================================================
log = logging.getLogger("RevitElement")

# Classes
# =========================================================================


class RevitElement:
    def __init__(self, doc, view, element):
        self.doc = doc
        self.view = view
        self.element = element

    @property
    def id(self):
        """Return integer id value or None."""
        return self.element.Id.Value if self.element else None

    @property
    def category(self):
        return self.element.Category.Name if self.element and self.element.Category else None

    def get_param(self, param_name, as_type=None, unit=None):
        if not self.element:
            return None

        p = self.element.LookupParameter(param_name)
        if not p:
            return None

        st = p.StorageType
        try:
            if as_type == "string" or (as_type is None and st == StorageType.String):
                s = p.AsString()
                return s if s is not None else p.AsValueString()
            if as_type == "int" or (as_type is None and st == StorageType.Integer):
                return p.AsInteger()
            if as_type == "double" or (as_type is None and st == StorageType.Double):
                val = p.AsDouble()
                if val is None:
                    return None
                if unit:
                    val = UnitUtils.ConvertFromInternalUnits(val, unit)
                return float(val)
            if as_type == "elementid" or (as_type is None and st == StorageType.ElementId):
                eid = p.AsElementId()
                return eid if isinstance(eid, ElementId) else None
        except Exception as ex:
            log.debug("get_param error for %s on %s: %s",
                      param_name, self.id, ex)
            return None

        # fallback: try value string
        try:
            return p.AsValueString()
        except Exception:
            return None

    def set_param(self, param_name, value):
        # Deterministic setter that follows the parameter StorageType. Returns True on success, False otherwise.
        if not self.element:
            return False

        p = self.element.LookupParameter(param_name)
        if not p:
            log.debug("Parameter '%s' not found on element %s",
                      param_name, self.id)
            return False

        if p.IsReadOnly:
            log.debug("Parameter '%s' is read-only on element %s",
                      param_name, self.id)
            return False

        st = p.StorageType
        try:
            if st == StorageType.String:
                p.Set(str(value))
                return True
            elif st == StorageType.Integer:
                p.Set(int(value))
                return True
            elif st == StorageType.Double:
                p.Set(float(value))
                return True
            elif st == StorageType.ElementId:
                if isinstance(value, ElementId):
                    p.Set(value)
                    return True
                # accept int id too
                if isinstance(value, int):
                    p.Set(ElementId(value))
                    return True
                log.debug(
                    "Value for ElementId param '%s' not ElementId or int", param_name)
                return False
        except Exception as ex:
            log.debug("set_param error for %s on %s: %s",
                      param_name, self.id, ex)
            return False

        log.debug("Unsupported storage type for '%s' on %s",
                  param_name, self.id)
        return False

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
            if hasattr(el, "element") and el.element:
                ids.Add(el.element.Id)
            elif isinstance(el, ElementId):
                ids.Add(el)
            elif hasattr(el, "Id"):
                ids.Add(el.Id)
        if ids.Count > 0:
            uidoc.Selection.SetElementIds(ids)
