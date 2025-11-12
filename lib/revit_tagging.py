# -*- coding: utf-8 -*-
############################################################################
# Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.
#
# This code and associated documentation files may not be copied, modified,
# distributed, or used in any form without the prior written permission of 
# the copyright holder.
############################################################################

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import UnitTypeId
from pyrevit import revit, forms, DB
from Autodesk.Revit.UI import UIDocument
from Autodesk.Revit.ApplicationServices import Application
from enum import Enum
import re

# Variables
# =======================================================================
app   = __revit__.Application           #type: Application
uidoc = __revit__.ActiveUIDocument      #type: UIDocument

#Class
# =======================================================================
class RevitXYZ(object):
    def __init__(self, element):
        self.element        = element
        self.loc            = getattr(element, "Location", None)
        self.curve          = getattr(self.loc, "Curve", None) if self.loc else None
        self.doc            = revit.doc
        self.view           = revit.active_view

    def start_point(self):
        if self.curve:
            return self.curve.GetEndPoint(0)
        return None

    def end_point(self):
        if self.curve:
            return self.curve.GetEndPoint(1)
        return None

    def midpoint(self):
        if self.curve:
            return self.curve.Evaluate(0.5, True)
        return None

    def point_at(self, param=0.25):
        if self.curve:
            return self.curve.Evaluate(param, True)
        return None
    
    def point_start(self, loc=("start", "end") point=("x", "z", "y")):
        if isinstance(loc, int):
            idx = 0 if loc == 0 else 1
        else:
            s = str(loc).strip().lower()
            if s in ("start", "s" "0"):
                idx = 0
            


class RevitTagging:
    def __init__(self, element):
        self.element    = element
        self.doc        = revit.doc
        self.view       = revit.active_view
        self.tag_syms   = (DB.FilteredElementCollector(self.doc)
                                .OfClass(DB.FamilySymbol)
                                .OfCategory(DB.BuiltInCategory.OST_FabricationDuctworkTags)
                                .ToElements())

    def get_label(self, name_contains):
        tag = name_contains.lower()
        for ts in self.tag_syms:
            fam = getattr(ts, "Family", None)
            fam_name = fam.Name if fam else ""
            ts_name = getattr(ts, "Name", "")
            pool = (fam_name + " " + ts_name).lower()
            if tag in pool:
                return ts
        raise LookupError("No label found with: " + name_contains)

    def already_tagged(self, elem, tag_fam_name):
        existing = (DB.FilteredElementCollector(self.doc, self.view.Id)
                    .OfClass(DB.IndependentTag)
                    .ToElements())
        for itag in existing:
            try:
                ref = itag.GetTaggedLocalElement()
            except:
                ref = None
            if ref and ref.Id == elem.Id:
                famname = itag.GetType().FamilyName
                if famname == tag_fam_name:
                    return True
        return False

    def place(self, element, tag_symbol, point_xyz):
        ref = DB.Reference(element)
        tag = DB.IndependentTag.Create(
            self.doc,
            self.view.Id,
            ref,
            False,
            DB.TagMode.TM_ADDBY_CATEGORY,
            DB.TagOrientation.Horizontal,
            point_xyz
        )
        if tag_symbol and tag_symbol.Id:
            tag.ChangeTypeId(tag_symbol.Id)
        return tag