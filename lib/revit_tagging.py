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
from revit_xyz import RevitXYZ
from revit_duct import RevitDuct
import re

# Variables
# =======================================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument


# Classes
# =======================================================================
class TagConfig(object):
    def __init__(self, names, tags, predicate=None, location_func=None):
        # names: tuple/list of accepted family names (lowercase)
        # tags: list of (tag, x_loc, z_offset)
        # predicate: callable(RevitDuct) -> bool
        # location_func: callable(RevitDuct, x_loc, z_offset) -> DB.XYZ or None
        self.names = tuple(n.strip().lower() for n in names)
        self.tags = tags
        self.predicate = predicate if predicate else (lambda d: True)
        self.location_func = location_func

    def matches(self, fam_name):
        return fam_name in self.names


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
                try:
                    tag_type = self.doc.GetElement(itag.GetTypeId())
                    if tag_type and hasattr(tag_type, 'Family'):
                        famname = tag_type.Family.Name
                        if famname == tag_fam_name:
                            return True
                except Exception:
                    continue
        return False

    def place_tag(self, element_or_ref, tag_symbol=None, point_xyz=None):
        """
        Create an IndependentTag attached to element or reference.
        - element_or_ref: either a Revit Element or a Reference (face/element)
        - tag_symbol: optional FamilySymbol to set type (pass symbol object)
        - point_xyz: XYZ insertion point
        Note: caller must open a Transaction before calling this method.
        """
        from Autodesk.Revit.DB import ElementId

        if element_or_ref is None:
            raise ValueError("element_or_ref is required")

        # If caller passed an element wrapper (e.g. RevitDuct), accept .element
        el_or_ref = getattr(element_or_ref, "element", element_or_ref)

        # If el_or_ref is a Reference already, use it; otherwise build Reference(element)
        if isinstance(el_or_ref, Reference):
            ref = el_or_ref
        else:
            ref = Reference(el_or_ref)

        tag = IndependentTag.Create(
            self.doc,
            self.view.Id,
            ref,
            False,
            TagMode.TM_ADDBY_CATEGORY,
            TagOrientation.Horizontal,
            point_xyz,
        )

        if tag_symbol is not None and getattr(tag_symbol, "Id", None):
            # Accept either ElementId or FamilySymbol
            new_id = tag_symbol.Id if hasattr(tag_symbol, "Id") else tag_symbol
            if isinstance(new_id, ElementId):
                tag.ChangeTypeId(new_id)

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
        # Use a list for mutability: [face, ndot, dist, centroid]
        best = [None, 1.0, float("inf"), None]

        def score_face(face, transform):
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
                    ab = b - a
                    ac = c - a
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
                dist = (
                    centroid.DistanceTo(prefer_point)
                    if prefer_point is not None
                    else 0.0
                )
                # choose face with minimal ndot; tie-breaker is smaller distance
                if ndot < best[1] or (abs(ndot - best[1]) < 1e-6 and dist < best[2]):
                    best[0] = face
                    best[1] = ndot
                    best[2] = dist
                    best[3] = centroid
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
