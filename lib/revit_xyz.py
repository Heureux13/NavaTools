# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import UnitTypeId
from pyrevit import revit, script, forms, DB
from Autodesk.Revit.UI import UIDocument
from Autodesk.Revit.ApplicationServices import Application
from enum import Enum
import logging
import math
import re

# Variables
# ==================================================s   
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
view = revit.active_view
output = script.get_output()

# Logging
log = logging.getLogger("RevitDuct")

# Constants
TOL = 1e-6

# Class logic
# ==================================================
class RevitXYZ(object):
    def __init__(self, element):
        self.element = element
        self.loc = getattr(element, "Location", None)
        self.curve = getattr(self.loc, "Curve", None) if self.loc else None
        self.doc = revit.doc
        self.view = revit.active_view

    def start_point(self):
        if self.curve:
            return self.curve.GetEndPoint(0)
        return None

    def end_point(self):
        if self.curve:
            return self.curve.GetEndPoint(1)
        return None

    def mid_point(self):
        if self.curve:
            return self.curve.Evaluate(0.5, True)
        return None

    def point_at(self, param=0.25):
        if self.curve:
            t = max(0.0, min(1.0, float(param)))
            return self.curve.Evaluate(t, True)
        return None

    def straight_joint_degree(self):
        """Returns the angle in degrees between the duct and the horizontal (XY) plane."""
        start = self.start_point()
        end = self.end_point()
        if not start or not end:
            return None

        dx = end.X - start.X
        dy = end.Y - start.Y
        dz = end.Z - start.Z

        horizontal_length = math.sqrt(dx**2 + dy**2)
        if horizontal_length == 0:
            return 90.0 if dz != 0 else 0.0

        angle_rad = math.atan2(dz, horizontal_length)
        angle_deg = math.degrees(angle_rad)
        return round(angle_deg, 2)

    def true_length(self):
        """Returns the true 3D length of the duct."""
        start = self.start_point()
        end = self.end_point()
        if not start or not end:
            return None

        dx = end.X - start.X
        dy = end.Y - start.Y
        dz = end.Z - start.Z

        length = math.sqrt(dx**2 + dy**2 + dz**2)
        return round(length, 2)
    
        # Geometry helpers
    def _get_geometry_solids(self):
        """Return a list of Solid objects found in the element's geometry."""
        opts = Options()
        try:
            geom = self.element.get_Geometry(opts)
        except Exception:
            return []

        solids = []
        for g in geom:
            try:
                if isinstance(g, Solid) and g.Volume > 0:
                    solids.append(g)
                elif isinstance(g, GeometryInstance):
                    inst = g.GetInstanceGeometry()
                    for ig in inst:
                        if isinstance(ig, Solid) and ig.Volume > 0:
                            solids.append(ig)
            except Exception:
                continue
        return solids

    def _face_centroid_and_normal(self, face):
        """Compute an approximate centroid and (unnormalized) normal vector for a Face using triangulation.

        Returns (centroid_xyz, normal_xyz) where centroid_xyz and normal_xyz are Autodesk.Revit.DB.XYZ
        or (None, None) if computation fails.
        """
        try:
            mesh = face.Triangulate()
            verts = list(mesh.Vertices)
            inds = list(mesh.TriangleIndices)
            if not verts or not inds:
                return (None, None)

            total_area = 0.0
            cx = cy = cz = 0.0
            nx = ny = nz = 0.0
            for i in range(0, len(inds), 3):
                try:
                    v0 = verts[inds[i]]
                    v1 = verts[inds[i+1]]
                    v2 = verts[inds[i+2]]
                except Exception:
                    continue

                ux = v1.X - v0.X
                uy = v1.Y - v0.Y
                uz = v1.Z - v0.Z

                vx = v2.X - v0.X
                vy = v2.Y - v0.Y
                vz = v2.Z - v0.Z

                # cross = u x v
                cxr = uy * vz - uz * vy
                cyr = uz * vx - ux * vz
                czr = ux * vy - uy * vx

                # triangle area
                tri_area = 0.5 * ((cxr * cxr + cyr * cyr + czr * czr) ** 0.5)
                if tri_area == 0:
                    continue

                # triangle centroid
                tcx = (v0.X + v1.X + v2.X) / 3.0
                tcy = (v0.Y + v1.Y + v2.Y) / 3.0
                tcz = (v0.Z + v1.Z + v2.Z) / 3.0

                cx += tcx * tri_area
                cy += tcy * tri_area
                cz += tcz * tri_area

                nx += cxr * tri_area
                ny += cyr * tri_area
                nz += czr * tri_area

                total_area += tri_area

            if total_area == 0:
                return (None, None)

            centroid = XYZ(cx / total_area, cy / total_area, cz / total_area)
            normal = XYZ(nx, ny, nz)
            return (centroid, normal)
        except Exception:
            return (None, None)

    def faces(self):
        """Return a list of Face objects for the element's solids."""
        faces = []
        solids = self._get_geometry_solids()
        for s in solids:
            try:
                for f in s.Faces:
                    faces.append(f)
            except Exception:
                continue
        return faces

    def faces_info(self):
        """Return a list of dicts describing each face: area, centroid, normal, and the Face object."""
        info = []
        for f in self.faces():
            try:
                area = getattr(f, 'Area', None)
                centroid, normal = self._face_centroid_and_normal(f)
                info.append({
                    'face': f,
                    'area': round(area, 4) if isinstance(area, (int, float)) else None,
                    'centroid': centroid,
                    'normal': normal,
                })
            except Exception:
                continue
        return info

    def get_face_reference_and_tag_point(self, offset_ft=0.1, prefer_largest=True, preferred_direction=None):
        """Return (face, reference, tag_point) for tagging.

        - offset_ft: distance (feet) to offset from the face along its outward normal.
        - prefer_largest: if True choose the largest face by area, otherwise choose by preferred_direction.
        - preferred_direction: an XYZ direction to prefer (pass when prefer_largest is False).

        Returns (face, reference, XYZ) or (None, None, None) if no usable face/reference found.
        """
        infos = self.faces_info()
        if not infos:
            return (None, None, None)

        chosen = None
        # If caller provided a direction and asked not to prefer largest, pick by alignment
        if preferred_direction is not None and not prefer_largest:
            pd = preferred_direction
            pd_mag = (pd.X * pd.X + pd.Y * pd.Y + pd.Z * pd.Z) ** 0.5
            if pd_mag != 0:
                pd = XYZ(pd.X / pd_mag, pd.Y / pd_mag, pd.Z / pd_mag)
                best_dot = -1.0
                for info in infos:
                    n = info.get('normal')
                    if not n:
                        continue
                    mag = (n.X * n.X + n.Y * n.Y + n.Z * n.Z) ** 0.5
                    if mag == 0:
                        continue
                    nu = XYZ(n.X / mag, n.Y / mag, n.Z / mag)
                    dot = abs(nu.X * pd.X + nu.Y * pd.Y + nu.Z * pd.Z)
                    if dot > best_dot:
                        best_dot = dot
                        chosen = info

        if chosen is None:
            # fallback: choose largest face by area
            infos_sorted = sorted(infos, key=lambda i: (i.get('area') or 0.0), reverse=True)
            chosen = infos_sorted[0] if infos_sorted else None

        if not chosen:
            return (None, None, None)

        face = chosen.get('face')
        centroid = chosen.get('centroid')
        normal = chosen.get('normal')
        if face is None or centroid is None or normal is None:
            return (face, None, None)

        # normalize normal
        mag = (normal.X * normal.X + normal.Y * normal.Y + normal.Z * normal.Z) ** 0.5
        if mag == 0:
            return (face, None, centroid)
        nu = XYZ(normal.X / mag, normal.Y / mag, normal.Z / mag)

        # compute tag point offset along normal
        tag_pt = XYZ(centroid.X + nu.X * float(offset_ft),
                     centroid.Y + nu.Y * float(offset_ft),
                     centroid.Z + nu.Z * float(offset_ft))

        # try to get a Reference for the face (may not exist in some contexts)
        ref = None
        try:
            if hasattr(face, 'Reference'):
                ref = face.Reference
            else:
                # some face objects expose GetReference
                getref = getattr(face, 'GetReference', None)
                if callable(getref):
                    ref = getref()
        except Exception:
            ref = None

        return (face, ref, tag_pt)

    def find_inlet_connector(self, elem=None):
        """Find the inlet connector for the element (or self.element if elem not provided)."""
        if elem is None:
            elem = self.element
        try:
            conns = list(elem.MEPModel.ConnectorManager.Connectors)
        except Exception:
            # older/newer API iterator variations
            try:
                conns = list(elem.MEPModel.ConnectorManager.Connectors.ForwardIterator())
            except Exception:
                return None
        # prefer explicit flow direction
        for c in conns:
            try:
                fd = getattr(c, "FlowDirection", None) or getattr(c, "Direction", None)
            except Exception:
                fd = None
            if fd is not None:
                # treat FlowDirectionType.In as inlet (adjust if your naming differs)
                if str(fd).upper().find("IN") >= 0 or str(fd).upper().find("INBOUND") >= 0:
                    return c
                # if Out means flow out of this connector, and you treat inlet as where air enters element
                if str(fd).upper().find("OUT") >= 0:
                    # depending on convention you might return c here
                    pass
        # fallback: prefer connector attached to equipment or supply system
        for c in conns:
            try:
                sys = c.GetConnectedSystem()
                if sys is not None and getattr(sys, "SystemType", None) is not None:
                    # check common supply types heuristically; adapt to your shop
                    if "Supply" in str(sys.SystemType):
                        return c
            except Exception:
                pass
        # final fallback: return nearest connector to element bounding box min corner
        if conns:
            try:
                bb = elem.get_BoundingBox(None)
                anchor = (bb.Min + bb.Max) * 0.5 if bb is not None else conns[0].Origin
            except Exception:
                anchor = conns[0].Origin
            return min(conns, key=lambda c: (c.Origin - anchor).GetLength())
        return None

    def face_perimeter_points(face, transform=None):
        pts = []
        try:
            loops = face.GetEdgesAsCurveLoops()
            for loop in loops:
                for c in loop:
                    p0 = c.GetEndPoint(0)
                    if transform:
                        p0 = transform.OfPoint(p0)
                    pts.append(p0)
            # remove duplicates (close loops)
            unique = []
            for p in pts:
                if not any((p - q).GetLength() < TOL for q in unique):
                    unique.append(p)
            return unique
        except Exception:
            # fallback to triangulation perimeter (may include interior verts)
            try:
                tris = face.Triangulate()
                verts = [tris.get_Vertex(i) for i in range(tris.NumberOfVertices)]
                # keep unique
                unique = []
                for p in verts:
                    if not any((p - q).GetLength() < TOL for q in unique):
                        unique.append(p)
                return unique
            except Exception:
                return []

    def pick_corner_lowest_x(corners_world, active_view):
        # compute view axes
        try:
            vdir = active_view.ViewDirection
        except Exception:
            vdir = active_view.GetOrientation().ForwardDirection
        view_up = getattr(active_view, "UpDirection", None) or active_view.GetOrientation().UpDirection
        view_right = view_up.CrossProduct(vdir).Normalize()
        view_up = view_up.Normalize()
        # project, find min X
        proj = [(p.DotProduct(view_right), p.DotProduct(view_up), p) for p in corners_world]
        min_x = min(p[0] for p in proj)
        # choose point with minimum X; if multiple, choose min Y among them
        candidates = [p for p in proj if abs(p[0] - min_x) < 1e-9]
        chosen = min(candidates, key=lambda p: p[1])[2]
        return chosen

    def face_for_connector(self, elem, connector):
        # pick face whose perimeter contains or is nearest to connector origin
        try:
            opts = Options()
            geom = elem.get_Geometry(opts)
        except Exception:
            geom = None
        faces = []
        if geom is None:
            return None
        def collect(g, xform = None):
            from Autodesk.Revit.DB import Solid, GeometryInstance, Transform
            if g is None:
                return
            if isinstance(g, Solid):
                for f in g.Faces:
                    faces.append((f, xform))
            elif isinstance(g, GeometryInstance):
                inst_xform = g.Transform
                inst_geom = g.GetInstanceGeometry()
                if inst_geom:
                    for gi in inst_geom:
                        collect(gi, inst_xform * (xform or Transform.Identity))
            else:
                try:
                    for gg in g:
                        collect(gg, xform)
                except Exception:
                    pass
        for g in geom:
            collect(g, None)
        # choose face nearest connector origin
        if not faces:
            return None, None
        conn_pt = connector.Origin
        best = min(faces, key=lambda fr: self.nearest_point_on_face(fr[0], fr[1], conn_pt).GetLength())
        return best  # (face, transform)

    def nearest_point_on_face(self, face, transform, pt):
        # project pt onto face (world -> face local if transform present)
        try:
            if transform:
                # convert pt to local
                inv = transform.Inverse
                local_pt = inv.OfPoint(pt)
                proj = face.Project(local_pt)
                return transform.OfPoint(proj.XYZPoint)
            else:
                proj = face.Project(pt)
                return proj.XYZPoint
        except Exception:
            # distance fallback
            return pt

    def inlet_corner_insertion(self, elem=None, active_view=None, tag_offset=0.1):
        # Combined flow: determine inlet corner (lowest X) and return insertion point offset by normal
        if elem is None:
            elem = self.element
        if active_view is None:
            active_view = self.view
        conn = self.find_inlet_connector(elem)
        if conn is None:
            return None, "No connector found"
        face_transform_pair = self.face_for_connector(elem, conn)
        if face_transform_pair is None:
            return None, "No face found"
        face, transform = face_transform_pair
        pts = self.face_perimeter_points(face, transform)
        if not pts:
            return None, "No perimeter points"
        corner = self.pick_corner_lowest_x(pts, active_view)
        # compute face normal (approx via triangle)
        try:
            tris = face.Triangulate()
            a = tris.get_Vertex(0); b = tris.get_Vertex(1); c = tris.get_Vertex(2)
            normal = (b - a).CrossProduct(c - a).Normalize()
            if transform:
                normal = transform.OfVector(normal).Normalize()
        except Exception:
            # fallback: compute via two edges from perimeter
            normal = XYZ(0,0,1)
        insertion = corner + normal * tag_offset
        return insertion, None

    # Usage example:
    # insertion, err = inlet_corner_insertion(my_duct_element, doc.ActiveView, tag_offset=0.1)
    # if insertion:
    #     # create/move tag here (ElementTransformUtils.MoveElement or IndependentTag creation)
    #     pass
