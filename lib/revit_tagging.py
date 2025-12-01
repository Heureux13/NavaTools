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
import re

# Variables
# =======================================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument


# Functions
# =======================================================================
def get_revit_year(app):
    """Extract Revit year from Application.VersionName."""
    name = app.VersionName
    for n in name.split():
        if n.isdigit():
            return int(n)
    return None


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

        tags = list(
            FilteredElementCollector(self.doc, self.view.Id)
            .OfClass(IndependentTag)
            .ToElements()
        )

        revit_year = get_revit_year(app)

        for itag in tags:
            try:
                tagged_elem_id = None

                # Revit 2026+ uses GetTaggedLocalElementIds() method
                if revit_year and revit_year >= 2026:
                    tagged_elem_ids = itag.GetTaggedLocalElementIds()
                    if not tagged_elem_ids:
                        continue
                    # Check if any of the tagged element IDs match our element
                    for tid in tagged_elem_ids:
                        if tid == elem.Id:
                            tagged_elem_id = tid
                            break
                # Revit 2022-2025 uses TaggedLocalElementId property
                else:
                    tagged_elem_id = itag.TaggedLocalElementId

                if tagged_elem_id and tagged_elem_id == elem.Id:
                    # Get the tag's FamilySymbol to check family name
                    tag_type_id = itag.GetTypeId()
                    tag_type = self.doc.GetElement(tag_type_id)
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

    @staticmethod
    def midpoint_location(d, x_loc, z_offset):
        loc = d.element.Location
        if hasattr(loc, "Curve") and loc.Curve:
            pt = loc.Curve.Evaluate(x_loc, True)
            return DB.XYZ(pt.X, pt.Y, pt.Z + z_offset)
        # fallback to bbox center
        v = getattr(d, "view", None) or revit.active_view
        bbox = d.element.get_BoundingBox(v) if v else None
        if bbox:
            center = (bbox.Min + bbox.Max) / 2.0
            return DB.XYZ(center.X, center.Y, center.Z + z_offset)
        return None

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

    def get_tag_point_on_face(
        self, offset_ft=0.1, prefer_largest=True, preferred_direction=None
    ):
        """Return a (face, point_xyz) suitable for placing a tag.

        - offset_ft: distance in feet to offset the tag point along the face normal so the tag is readable.
        - prefer_largest: if True pick the largest face by area; otherwise pick the face whose normal
          is closest to the preferred_direction (an XYZ) if provided, otherwise largest.

        Returns (face, XYZ) or (None, None) if no usable face found.
        """
        try:
            rxyz = RevitXYZ(self.element)
            infos = rxyz.faces_info()
            if not infos:
                return (None, None)

            # pick face
            chosen = None
            if preferred_direction is not None and not prefer_largest:
                # choose face whose normal best aligns with preferred_direction
                best = None
                best_dot = -1.0
                pd = preferred_direction
                pd_mag = (pd.X**2 + pd.Y**2 + pd.Z**2) ** 0.5
                if pd_mag == 0:
                    pd = None
                else:
                    pd = XYZ(pd.X / pd_mag, pd.Y / pd_mag, pd.Z / pd_mag)

                if pd is not None:
                    for info in infos:
                        n = info.get("normal")
                        if n is None:
                            continue
                        mag = (n.X**2 + n.Y**2 + n.Z**2) ** 0.5
                        if mag == 0:
                            continue
                        nu = XYZ(n.X / mag, n.Y / mag, n.Z / mag)
                        dot = abs(nu.X * pd.X + nu.Y * pd.Y + nu.Z * pd.Z)
                        if dot > best_dot:
                            best_dot = dot
                            best = info
                    chosen = best

            if chosen is None:
                # fallback: largest area
                infos_sorted = sorted(
                    infos, key=lambda i: (i.get("area") or 0.0), reverse=True
                )
                chosen = infos_sorted[0] if infos_sorted else None

            if not chosen:
                return (None, None)

            face = chosen.get("face")
            centroid = chosen.get("centroid")
            normal = chosen.get("normal")
            if centroid is None or normal is None:
                return (face, None)

            # normalize normal
            mag = (normal.X**2 + normal.Y**2 + normal.Z**2) ** 0.5
            if mag == 0:
                return (face, centroid)
            nu = XYZ(normal.X / mag, normal.Y / mag, normal.Z / mag)

            # compute offset point
            px = centroid.X + nu.X * float(offset_ft)
            py = centroid.Y + nu.Y * float(offset_ft)
            pz = centroid.Z + nu.Z * float(offset_ft)
            tag_point = XYZ(px, py, pz)
            return (face, tag_point)
        except Exception:
            return (None, None)
