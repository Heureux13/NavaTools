# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    FamilySymbol,
    IndependentTag,
    Reference,
    TagMode,
    TagOrientation,
    ElementId,
)
from Autodesk.Revit.DB import UnitTypeId
from Autodesk.Revit.UI import UIDocument
from pyrevit import revit, forms, DB
from Autodesk.Revit.ApplicationServices import Application
from pyrevit import revit, forms, DB
from enum import Enum
import re

# Variables
# =======================================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument


# Classes
# =======================================================================
class RevitTagging:
    """
    Helpers for finding tag family symbols and placing IndependentTag on elements.
    """

    def __init__(self, doc=None, view=None):
        self.doc = doc or revit.doc
        self.view = view or revit.active_view
        # Cache tag family symbols for fabrication ductwork tags
        self.tag_syms = (
            FilteredElementCollector(self.doc)
            .OfClass(FamilySymbol)
            .OfCategory(DB.BuiltInCategory.OST_FabricationDuctworkTags)
            .ToElements()
        )

    def get_label(self, name_contains):
        """Return the first FamilySymbol whose family or type name contains the substring."""
        if not name_contains:
            raise ValueError("name_contains must be a non-empty string")
        needle = name_contains.lower()
        for ts in self.tag_syms:
            fam = getattr(ts, "Family", None)
            fam_name = fam.Name if fam is not None else ""
            ts_name = getattr(ts, "Name", "") or ""
            pool = (fam_name + " " + ts_name).lower()
            if needle in pool:
                return ts
        raise LookupError("No label found with: " + name_contains)

    def already_tagged(self, elem, tag_fam_name):
        """
        Check whether the element already has a tag of the specified family name
        in the current view. Returns True/False.
        """
        if elem is None:
            return False

        tags = (
            FilteredElementCollector(self.doc, self.view.Id)
            .OfClass(IndependentTag)
            .ToElements()
        )
        for itag in tags:
            # try to resolve the tagged element reference safely
            try:
                tagged_el = itag.GetTaggedLocalElement()
            except Exception:
                # fallback: try TaggedLocalElementId or other APIs depending on Revit version
                try:
                    eid = itag.TaggedLocalElementId
                    tagged_el = self.doc.GetElement(eid) if eid else None
                except Exception:
                    tagged_el = None

            if tagged_el is None:
                continue

            if tagged_el.Id == elem.Id:
                famname = itag.GetType().FamilyName if itag.GetType() is not None else ""
                if famname == tag_fam_name:
                    return True
        return False

    def place(self, element, tag_symbol, point_xyz):
        """
        Place an independent tag on the element at the given XYZ point.
        - element: the Revit element to tag
        - tag_symbol: FamilySymbol for the tag type (can be None to use default)
        - point_xyz: XYZ location for the tag head
        Returns the created IndependentTag instance.
        """
        if element is None:
            raise ValueError("element is required")

        ref = Reference(element)
        tag = IndependentTag.Create(
            self.doc,
            self.view.Id,
            ref,
            False,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            point_xyz,
        )
        # change type if provided
        if tag_symbol is not None and getattr(tag_symbol, "Id", None):
            tag.ChangeTypeId(tag_symbol.Id)
        return tag
