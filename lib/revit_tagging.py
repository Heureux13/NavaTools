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
from Autodesk.Revit.UI import UIDocument
from pyrevit import revit, forms, DB
from Autodesk.Revit.ApplicationServices import Application
from enum import Enum
import re

# Variables
# =======================================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument


# Classes
# =======================================================================
class RevitTagging:
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

    def get_face_facing_view(self, element, prefer_point=None):
        """
        Return (Reference, centroid_XYZ) for the face of `element` that best faces
        the current view (self.view). Optionally prefer faces near `prefer_point`.
        Returns (None, None) if no suitable face found.

        Notes:
        - Uses Options.ComputeReferences = True so the returned Reference can be used
        directly with IndependentTag.Create.
        - This method handles GeometryInstance transforms.
        """
        from Autodesk.Revit.DB import Options, GeometryInstance, Solid, XYZ

        if element is None:
            return None, None

        opt = Options()
        opt.DetailLevel = getattr(self.view, "DetailLevel", None)
        opt.ComputeReferences = True

        try:
            geom = element.get_Geometry(opt)
        except Exception:
            return None, None

        world_dir = self.view.ViewDirection  # vector from view to model
        best = (None, 1.0, float("inf"), None)  # (face, dot_with_view_dir, dist_to_pref, centroid)

        def score_face(face, transform):
            nonlocal best
            try:
                tri = face.Triangulate()
                verts = list(tri.Vertices)
                if not verts:
                    return
                # centroid (in local coords); transform if needed
                cx = sum(v.X for v in verts) / len(verts)
                cy = sum(v.Y for v in verts) / len(verts)
                cz = sum(v.Z for v in verts) / len(verts)
                centroid = XYZ(cx, cy, cz)
                if transform is not None:
                    centroid = transform.OfPoint(centroid)

                # approximate normal using first triangle
                try:
                    a, b, c = verts[0], verts[1], verts[2]
                    ab = b - a; ac = c - a
                    n = ab.CrossProduct(ac)
                    nlen = n.GetLength()
                    if nlen == 0:
                        ndot = 0.0
                    else:
                        ndot = n.Normalize().DotProduct(world_dir)
                except Exception:
                    ndot = 0.0

                # prefer faces that face the view (ndot should be negative);
                # smaller ndot (more negative) is better.
                dist = centroid.DistanceTo(prefer_point) if prefer_point is not None else 0.0
                # choose face with minimal ndot; tie-breaker is smaller distance
                if ndot < best[1] or (abs(ndot - best[1]) < 1e-6 and dist < best[2]):
                    best = (face, ndot, dist, centroid)
            except Exception:
                return

        for g in geom:
            if isinstance(g, GeometryInstance):
                tr = g.Transform
                try:
                    inst_geo = g.GetInstanceGeometry()
                except Exception:
                    continue
                for sg in inst_geo:
                    if isinstance(sg, Solid) and sg.Volume > 0:
                        for f in sg.Faces:
                            score_face(f, tr)
            else:
                if isinstance(g, Solid) and g.Volume > 0:
                    for f in g.Faces:
                        score_face(f, None)

        face, ndot, dist, centroid = best
        if face is None:
            return None, None

        try:
            return face.Reference, centroid
        except Exception:
            return None, centroid