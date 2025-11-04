# duct_tagger.py
# Class-based duct tagging utilities with separate methods
# Copyright (c) 2025

from pyrevit import revit, DB
from Autodesk.Revit.DB import FabricationPart
import math


class DuctTagger:
    def __init__(self, doc, view):
        self.doc = doc
        self.view = view
        self.tag_syms = (DB.FilteredElementCollector(doc)
                         .OfClass(DB.FamilySymbol)
                         .OfCategory(DB.BuiltInCategory.OST_FabricationDuctworkTags)
                         .ToElements())

        self.down_nudge = 0.05
        self.left_nudge = 0.05

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

    def get_duct_corners(self, curve, width, height):
        p0, p1 = curve.GetEndPoint(0), curve.GetEndPoint(1)
        x_axis = (p1 - p0).Normalize()
        y_axis = self.view.RightDirection.Normalize()
        z_axis = self.view.UpDirection.Normalize()

        half_w = 0.5 * width
        half_h = 0.5 * height

        def corners_at(pt):
            return [
                pt + y_axis.Multiply(-half_w) +
                z_axis.Multiply(-half_h),  # bottom-left
                pt + y_axis.Multiply(half_w) +
                z_axis.Multiply(-half_h),  # bottom-right
                pt + y_axis.Multiply(-half_w) +
                z_axis.Multiply(half_h),  # top-left
                pt + y_axis.Multiply(half_w) +
                z_axis.Multiply(half_h),  # top-right
            ]

        return corners_at(p0), corners_at(p1)

    def tag_horizontal(self, tag_name="length"):
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

            # Orientation check: horizontal = u dominates
            p0, p1 = curve.GetEndPoint(0), curve.GetEndPoint(1)
            vec = p1 - p0
            u = vec.DotProduct(self.view.RightDirection)
            v = vec.DotProduct(self.view.UpDirection)

            if abs(v) > abs(u) * 0.2:
                continue

            bbox = duct.get_BoundingBox(
                self.view) or duct.get_BoundingBox(None)
            if not bbox:
                continue
            origin = (bbox.Min + bbox.Max) / 2.0

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

            placement = self.reconstruct_from_relative(umin, vmin, origin)

            tag = self.place(duct, len_tag, origin)
            tag.TagHeadPosition = placement
            tag.TagOrientation = DB.TagOrientation.Horizontal
            tag.HasLeader = False

            self.doc.Regenerate()

        print("Ducts found:", len(ducts))
        print("Tag symbols found:", len(self.tag_syms))

        t.Commit()

    def tag_vertical(self, tag_name="length"):
        ducts = (DB.FilteredElementCollector(self.doc, self.view.Id)
                 .OfCategory(DB.BuiltInCategory.OST_FabricationDuctwork)
                 .WhereElementIsNotElementType()
                 .ToElements())

        len_tag = self.get_label(tag_name)

        t = DB.Transaction(self.doc, "Tag Vertical Ducts")
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

            # Orientation check: vertical = v dominates
            p0, p1 = curve.GetEndPoint(0), curve.GetEndPoint(1)
            vec = p1 - p0
            u = vec.DotProduct(self.view.RightDirection)
            v = vec.DotProduct(self.view.UpDirection)

            if abs(u) > abs(v) * 0.2:
                continue

            bbox = duct.get_BoundingBox(
                self.view) or duct.get_BoundingBox(None)
            if not bbox:
                continue
            origin = (bbox.Min + bbox.Max) / 2.0

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

            # Bottom-right corner of duct bbox
            placement = self.reconstruct_from_relative(umax, vmin, origin)

            tag = self.place(duct, len_tag, origin)
            tag.TagHeadPosition = placement
            tag.TagOrientation = DB.TagOrientation.Horizontal
            tag.HasLeader = False

            self.doc.Regenerate()

        print("Ducts found:", len(ducts))
        print("Tag symbols found:", len(self.tag_syms))

        t.Commit()

    def tag_angled_down(self, tag_name="length"):
        ducts = (DB.FilteredElementCollector(self.doc, self.view.Id)
                 .OfCategory(DB.BuiltInCategory.OST_FabricationDuctwork)
                 .WhereElementIsNotElementType()
                 .ToElements())

        len_tag = self.get_label(tag_name)
        t = DB.Transaction(self.doc, "Tag Angled Down Ducts")
        t.Start()

        if not len_tag.IsActive:
            len_tag.Activate()
            self.doc.Regenerate()

        right = self.view.RightDirection.Normalize()
        up = self.view.UpDirection.Normalize()

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

            p0, p1 = curve.GetEndPoint(0), curve.GetEndPoint(1)
            dirv = (p1 - p0).Normalize()

            u = dirv.DotProduct(right)
            v = dirv.DotProduct(up)

            # angled down = u and v opposite signs
            if abs(u) < 1e-6 or abs(v) < 1e-6 or u * v >= 0:
                continue

            origin = (p0 + p1) / 2.0

            # Always pick the left endpoint
            ep0_u = (p0 - origin).DotProduct(right)
            ep1_u = (p1 - origin).DotProduct(right)
            anchor = p0 if ep0_u < ep1_u else p1

            # Placement = left endpoint + nudges only
            placement = (anchor
                         + right.Multiply(-self.left_nudge)
                         + up.Multiply(-self.down_nudge))

            print("Tagging duct:", duct.Id, "anchor:",
                  anchor, "placement:", placement)

            tag = self.place(duct, len_tag, origin)
            tag.TagHeadPosition = placement
            tag.TagOrientation = DB.TagOrientation.Horizontal
            tag.HasLeader = False

        self.doc.Regenerate()
        t.Commit()

    def tag_angled_up(self, tag_name="length"):
        ducts = (DB.FilteredElementCollector(self.doc, self.view.Id)
                 .OfCategory(DB.BuiltInCategory.OST_FabricationDuctwork)
                 .WhereElementIsNotElementType()
                 .ToElements())

        len_tag = self.get_label(tag_name)
        t = DB.Transaction(self.doc, "Tag Angled Up Ducts")
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

            # orientation check
            p0, p1 = curve.GetEndPoint(0), curve.GetEndPoint(1)
            vec = p1 - p0
            u = vec.DotProduct(self.view.RightDirection)
            v = vec.DotProduct(self.view.UpDirection)

            # skip horizontals and verticals
            if abs(v) <= abs(u) * 0.2:
                continue
            if abs(u) <= abs(v) * 0.2:
                continue

            # only keep upward diagonals u and v opposite sign
            if u * v >= 0:
                continue

            bbox = duct.get_BoundingBox(
                self.view) or duct.get_BoundingBox(None)
            if not bbox:
                continue
            origin = (bbox.Min + bbox.Max) / 2.0

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

            # place tag at top right for upward diagonals
            placement = self.reconstruct_from_relative(umax, vmax, origin)

            tag = self.place(duct, len_tag, origin)
            tag.TagHeadPosition = placement
            tag.TagOrientation = DB.TagOrientation.Horizontal
            tag.HasLeader = False

            self.doc.Regenerate()

        print("Ducts found:", len(ducts))
        print("Tag symbols found:", len(self.tag_syms))

        t.Commit()
