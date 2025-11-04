# -*- coding: utf-8 -*-
"""
=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
=========================================================================
"""

from pyrevit import revit, DB
from Autodesk.Revit.DB import FabricationPart


class TagDuct:
    def __init__(self, doc, view):
        self.doc = doc
        self.view = view
        self.tag_syms = (DB.FilteredElementCollector(doc)
                         .OfClass(DB.FamilySymbol)
                         .OfCategory(DB.BuiltInCategory.OST_FabricationDuctworkTags)
                         .ToElements())

    # ----------------- helpers -----------------
    def get_label(self, name_contains):
        tag = name_contains.lower()
        for ts in self.tag_syms:
            fam = getattr(ts, "Family", None)
            fam_name = fam.Name if fam else ""
            ts_name = getattr(ts, "Name", "")
            pool = (fam_name + " " + ts_name).lower()
            if tag in pool:
                return ts
        raise Exception("No label found with: " + name_contains)

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

    def project_relative(self, p, origin):
        v = p - origin
        return (v.DotProduct(self.view.RightDirection),
                v.DotProduct(self.view.UpDirection))

    def reconstruct_from_relative(self, u, v, origin):
        return (origin
                + self.view.RightDirection.Multiply(u)
                + self.view.UpDirection.Multiply(v))

    # ----------------- main routines -----------------
    def tag_horizontal(self, tag_name="length"):
        """Tag only horizontal ducts in the active view."""
        ducts = (DB.FilteredElementCollector(self.doc, self.view.Id)
                 .OfCategory(DB.BuiltInCategory.OST_FabricationDuctwork)
                 .WhereElementIsNotElementType()
                 .ToElements())

        len_tag = self.get_label(tag_name)

        t = DB.Transaction(self.doc, "Tag Horizontal Ducts")
        t.Start()

        if not len_tag.IsActive:
            len_tag.Activate()
            self.doc.Regenerate()

        for duct in ducts:
            if not isinstance(duct, FabricationPart):
                continue
            if self.already_tagged(duct, len_tag.Family.Name):
                continue

            loc = getattr(duct, "Location", None)
            if not isinstance(loc, DB.LocationCurve):
                continue

            curve = loc.Curve
            if not curve:
                continue

            # Orientation check using view projection
            p0, p1 = curve.GetEndPoint(0), curve.GetEndPoint(1)
            vec = p1 - p0
            u = vec.DotProduct(self.view.RightDirection)
            v = vec.DotProduct(self.view.UpDirection)

            # Only keep horizontals: u dominates, v is small
            if abs(v) > abs(u) * 0.2:
                continue

            bbox = duct.get_BoundingBox(
                self.view) or duct.get_BoundingBox(None)
            if not bbox:
                continue
            origin = (bbox.Min + bbox.Max) / 2.0

            # Project duct bbox into view space
            corners = [
                DB.XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z),
                DB.XYZ(bbox.Min.X, bbox.Min.Y, bbox.Max.Z),
                DB.XYZ(bbox.Min.X, bbox.Max.Y, bbox.Min.Z),
                DB.XYZ(bbox.Min.X, bbox.Max.Y, bbox.Max.Z),
                DB.XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z),
                DB.XYZ(bbox.Max.X, bbox.Min.Y, bbox.Max.Z),
                DB.XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z),
                DB.XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z),
            ]
            projected = [self.project_relative(c, origin) for c in corners]
            u_vals = [uv[0] for uv in projected]
            v_vals = [uv[1] for uv in projected]
            umin, umax = min(u_vals), max(u_vals)
            vmin, vmax = min(v_vals), max(v_vals)

            # Bottom-left corner of duct bbox
            chosen_u = umin
            chosen_v = vmin
            placement = self.reconstruct_from_relative(
                chosen_u, chosen_v, origin)

            # Create tag at dummy point, then force head
            tag = self.place(duct, len_tag, origin)
            tag.TagHeadPosition = placement
            tag.TagOrientation = DB.TagOrientation.Horizontal
            tag.HasLeader = False

            self.doc.Regenerate()

        t.Commit()
