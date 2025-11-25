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

    @staticmethod
    def normalize(v):
        """Normalize a 3D vector to unit length."""
        x, y, z = v
        m = math.sqrt(x*x + y*y + z*z)
        if m == 0:
            return (0.0, 0.0, 0.0)
        return (x/m, y/m, z/m)

    @staticmethod
    def dot(a, b):
        """Dot product of two 3D vectors."""
        return a[0]*b[0] + a[1]*b[1] + a[2]*b[2]

    @staticmethod
    def edge_diffs_whole_in(c0, W0, H0, c1, W1, H1, u_hat, v_hat, snap_tol=1/16.0):
        """
        Calculate edge-to-edge offsets for duct transitions.

        Parameters:
            c0, c1: Connector center positions (x, y, z) in inches
            W0, H0, W1, H1: Widths and heights in inches
            u_hat: Width axis unit vector (parallel to view on plan)
            v_hat: Height axis unit vector (perpendicular to u_hat on plan)
            snap_tol: Values below this snap to 0 before rounding (default 1/16")

        Returns:
            Dictionary with 'exact_in' and 'whole_in' offsets for left/right/top/bottom edges
        """
        # Check for None values
        if None in [W0, H0, W1, H1]:
            return None

        u_hat = RevitXYZ.normalize(u_hat)
        v_hat = RevitXYZ.normalize(v_hat)

        # Scalar projections of centers along u and v
        u0 = RevitXYZ.dot(c0, u_hat)
        u1 = RevitXYZ.dot(c1, u_hat)
        v0 = RevitXYZ.dot(c0, v_hat)
        v1 = RevitXYZ.dot(c1, v_hat)

        exact = {
            "left":   (u1 - W1/2.0) - (u0 - W0/2.0),
            "right":  (u1 + W1/2.0) - (u0 + W0/2.0),
            "top":    (v1 + H1/2.0) - (v0 + H0/2.0),
            "bottom": (v1 - H1/2.0) - (v0 - H0/2.0),
        }

        def snap_round(val):
            v2 = 0.0 if val < snap_tol else val
            return int(round(v2))

        whole = {k: snap_round(v) for k, v in exact.items()}
        return {"exact_in": exact, "whole_in": whole}

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

    def connector_elevation(self, connector_index):
        """Get Z elevation of a connector in feet."""
        connector = self.get_connector(connector_index)
        return connector.Origin.Z if connector else None

    def higher_connector_index(self):
        """Return the index (0 or 1) of the higher connector, or None if can't determine."""
        c0 = self.get_connector(0)
        c1 = self.get_connector(1)

        if not c0 or not c1:
            return None

        z0 = c0.Origin.Z
        z1 = c1.Origin.Z

        if abs(z1 - z0) < 1e-6:  # essentially equal elevation
            return None

        return 1 if z1 > z0 else 0

    def is_connector_higher(self, connector_index, than_index):
        """Check if connector at connector_index is higher than connector at than_index."""
        c1 = self.get_connector(connector_index)
        c2 = self.get_connector(than_index)

        if not c1 or not c2:
            return None

        return c1.Origin.Z > c2.Origin.Z

    @staticmethod
    def hv_offsets_inches(p1, p2):
        """
        p1, p2: (x, y, z) in inches
        returns whole-inch horizontal and vertical offsets
        """
        x1, y1, z1 = p1
        x2, y2, z2 = p2
        dx, dy, dz = x2 - x1, y2 - y1, z2 - z1
        H = math.sqrt(dx*dx + dy*dy)
        V = abs(dz)
        return {
            "horizontal_offset_in": int(round(H)),
            "vertical_offset_in": int(round(V)),
            "exact_horizontal_in": H,
            "exact_vertical_in": V
        }

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
            infos_sorted = sorted(infos, key=lambda i: (
                i.get('area') or 0.0), reverse=True)
            chosen = infos_sorted[0] if infos_sorted else None

        if not chosen:
            return (None, None, None)

        face = chosen.get('face')
        centroid = chosen.get('centroid')
        normal = chosen.get('normal')
        if face is None or centroid is None or normal is None:
            return (face, None, None)

        # normalize normal
        mag = (normal.X * normal.X + normal.Y *
               normal.Y + normal.Z * normal.Z) ** 0.5
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
                conns = list(
                    elem.MEPModel.ConnectorManager.Connectors.ForwardIterator())
            except Exception:
                return None
        # prefer explicit flow direction
        for c in conns:
            try:
                fd = getattr(c, "FlowDirection", None) or getattr(
                    c, "Direction", None)
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
                anchor = (bb.Min + bb.Max) * \
                    0.5 if bb is not None else conns[0].Origin
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
                verts = [tris.get_Vertex(i)
                         for i in range(tris.NumberOfVertices)]
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
        view_up = getattr(active_view, "UpDirection",
                          None) or active_view.GetOrientation().UpDirection
        view_right = view_up.CrossProduct(vdir).Normalize()
        view_up = view_up.Normalize()
        # project, find min X
        proj = [(p.DotProduct(view_right), p.DotProduct(view_up), p)
                for p in corners_world]
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

        def collect(g, xform=None):
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
            return None
        conn_pt = connector.Origin
        best = min(faces, key=lambda fr: self.nearest_point_on_face(
            fr[0], fr[1], conn_pt).GetLength())
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
            a = tris.get_Vertex(0)
            b = tris.get_Vertex(1)
            c = tris.get_Vertex(2)
            normal = (b - a).CrossProduct(c - a).Normalize()
            if transform:
                normal = transform.OfVector(normal).Normalize()
        except Exception:
            # fallback: compute via two edges from perimeter
            normal = XYZ(0, 0, 1)
        insertion = corner + normal * tag_offset
        return insertion, None

    # Usage example:
    # insertion, err = inlet_corner_insertion(my_duct_element, doc.ActiveView, tag_offset=0.1)
    # if insertion:
    #     # create/move tag here (ElementTransformUtils.MoveElement or IndependentTag creation)
    #     pass

    # =========================================
    # Transition / Reducer analysis (FOT/FOB/CL/FOS)
    # =========================================
    def _get_connectors(self, elem=None):
        if elem is None:
            elem = self.element
        # Try common connector access patterns across MEP and Fabrication
        for path in [
            "MEPModel.ConnectorManager.Connectors",
            "ConnectorManager.Connectors",
        ]:
            try:
                obj = elem
                for attr in path.split('.'):
                    obj = getattr(obj, attr)
                # Try to iterate
                try:
                    return list(obj)
                except Exception:
                    try:
                        it = obj.ForwardIterator()
                        items = []
                        while it.MoveNext():
                            items.append(it.Current)
                        return items
                    except Exception:
                        continue
            except Exception:
                continue
        return []

    def _end_faces_by_opposition(self, infos):
        """Pick two end faces whose normals are near-opposite and with large areas.

        Returns (info_in, info_out) as dicts from faces_info(). If multiple, picks the best pair.
        """
        if not infos:
            return (None, None)
        # Precompute unit normals
        normed = []
        for i, inf in enumerate(infos):
            n = inf.get('normal')
            if n is None:
                continue
            mag = (n.X*n.X + n.Y*n.Y + n.Z*n.Z) ** 0.5
            if mag < TOL:
                continue
            nu = XYZ(n.X/mag, n.Y/mag, n.Z/mag)
            normed.append((i, inf, nu))
        if not normed:
            return (None, None)
        # Find pair with most opposite normals and largest combined area
        best = None
        best_score = -1e9
        for i in range(len(normed)):
            for j in range(i+1, len(normed)):
                ii, info_i, ni = normed[i]
                jj, info_j, nj = normed[j]
                dp = ni.DotProduct(nj)
                # Opposition score: more negative dp is better (approach -1)
                opp = -dp
                area_sum = (info_i.get('area') or 0.0) + \
                    (info_j.get('area') or 0.0)
                score = opp*1000.0 + area_sum
                if score > best_score:
                    best_score = score
                    best = (info_i, info_j)
        if best is None:
            return (None, None)
        return best

    def _face_top_bottom_elev(self, face, transform=None):
        """Return (zmin, zmax, centroid) in world coords for a planar face perimeter."""
        if face is None:
            return (None, None, None)
        pts = self.face_perimeter_points(face, transform)
        if not pts:
            # fallback to triangulation vertices
            try:
                tris = face.Triangulate()
                pts = [tris.get_Vertex(i)
                       for i in range(tris.NumberOfVertices)]
            except Exception:
                return (None, None, None)
        zvals = [p.Z for p in pts]
        zmin = min(zvals)
        zmax = max(zvals)
        # centroid approx
        cx = sum(p.X for p in pts)/float(len(pts))
        cy = sum(p.Y for p in pts)/float(len(pts))
        cz = sum(p.Z for p in pts)/float(len(pts))
        return (zmin, zmax, XYZ(cx, cy, cz))

    def analyze_transition(self, active_view=None):
        """Analyze a transition/reducer element for FOT/FOB/CL and side offset (FOS).

        Returns a dict with keys:
        - vertical_class: 'FOT'|'FOB'|'CL'|None
        - vertical_shift_ft: signed float (+up, -down)
        - vertical_arrow: string like '↑2"' or '↓1-1/2"' or ''
        - side_class: 'FOS'|None
        - side_dir: 'left'|'right'|None
        - side_shift_ft: signed float (+right, -left)
        - side_arrow: '→2"'|'←1"' or ''
        - in_centroid, out_centroid: XYZ
        - debug: dict with supporting values
        """
        if active_view is None:
            active_view = self.view

        infos = self.faces_info()
        fi, fo = self._end_faces_by_opposition(infos)
        if fi is None or fo is None:
            return {'vertical_class': None, 'vertical_shift_ft': 0.0, 'vertical_arrow': '',
                    'side_class': None, 'side_dir': None, 'side_shift_ft': 0.0, 'side_arrow': '',
                    'in_centroid': None, 'out_centroid': None,
                    'debug': {'reason': 'no end faces', 'num_faces': len(infos)}}

        # If we can find inlet connector, use the nearest end face as inlet
        inlet = self.find_inlet_connector(self.element)
        if inlet is not None:
            # pick face whose centroid is nearest to inlet origin
            ci = fi.get('centroid')
            co = fo.get('centroid')
            if ci is not None and co is not None:
                di = (ci - inlet.Origin).GetLength()
                do = (co - inlet.Origin).GetLength()
                if do < di:
                    fi, fo = fo, fi

        # Pull top/bottom elevations and centroids
        zmin_i, zmax_i, ci = self._face_top_bottom_elev(fi.get('face'))
        zmin_o, zmax_o, co = self._face_top_bottom_elev(fo.get('face'))
        if None in (zmin_i, zmax_i, zmin_o, zmax_o, ci, co):
            return {'vertical_class': None, 'vertical_shift_ft': 0.0, 'vertical_arrow': '',
                    'side_class': None, 'side_dir': None, 'side_shift_ft': 0.0, 'side_arrow': '',
                    'in_centroid': ci, 'out_centroid': co, 'debug': {'reason': 'missing z bounds'}}

        # Vertical classification (tolerance: 0.01 ft ~ 1/8 inch)
        TOL_FT = 0.01
        top_diff = abs(zmax_o - zmax_i)
        bot_diff = abs(zmin_o - zmin_i)
        top_same = top_diff < TOL_FT
        bot_same = bot_diff < TOL_FT
        vertical_class = None
        if top_same and bot_same:
            # Both top and bottom equal = centerline reducer
            vertical_class = 'CL'
        elif top_same and not bot_same:
            vertical_class = 'FOT'
        elif bot_same and not top_same:
            vertical_class = 'FOB'
        elif abs(top_diff - bot_diff) < TOL_FT:
            # Top and bottom move by same amount = centerline (parallel offset)
            vertical_class = 'CL'

        # Vertical center shift (signed): +up, -down
        cz_i = 0.5*(zmax_i + zmin_i)
        cz_o = 0.5*(zmax_o + zmin_o)
        v_shift = cz_o - cz_i

        def _gcd(a, b):
            a = int(abs(a))
            b = int(abs(b))
            while b:
                a, b = b, a % b
            return a if a != 0 else 1

        def fmt_inches(ft_val):
            sign = 1.0 if ft_val >= 0 else -1.0
            inches = abs(ft_val) * 12.0
            # round to nearest 1/16"
            inc16 = round(inches * 16.0) / 16.0
            if abs(inc16) < 1e-4:
                return ''
            # build string like 2-3/4"
            whole = int(math.floor(inc16 + 1e-6))
            frac = inc16 - whole

            def frac_str(x):
                denom = 16
                num = int(round(x * denom))
                if num == 0:
                    return ''
                # reduce fraction
                g = _gcd(num, denom)
                return "{}/{}".format(num//g, denom//g)
            if whole > 0 and frac > 1e-6:
                s = "{}-{}\"".format(whole, frac_str(frac))
            elif whole > 0:
                s = "{}\"".format(whole)
            else:
                s = "{}\"".format(frac_str(frac))
            return (u"↑" + s) if sign > 0 else (u"↓" + s)

        v_arrow = fmt_inches(v_shift)

        # Side (horizontal) offset relative to flow direction
        axis = (co - ci)
        # project onto horizontal plane
        zunit = XYZ(0, 0, 1)
        axis_h = axis - zunit.Multiply(axis.DotProduct(zunit))
        try:
            ah_len = axis_h.GetLength()
        except Exception:
            ah_len = 0.0
        if ah_len < 1e-8:
            # vertical part; choose arbitrary right vector
            right = XYZ(1, 0, 0)
        else:
            axis_hu = XYZ(axis_h.X/ah_len, axis_h.Y/ah_len, axis_h.Z/ah_len)
            # right-hand rule: right = flow x up
            right = axis_hu.CrossProduct(zunit)
            rlen = right.GetLength()
            if rlen > 1e-8:
                right = XYZ(right.X/rlen, right.Y/rlen, right.Z/rlen)
            else:
                right = XYZ(1, 0, 0)
        lateral = (co - ci).DotProduct(right)
        side_class = 'FOS' if abs(lateral) > 1e-5 else None
        side_dir = 'right' if lateral > 0 else (
            'left' if lateral < 0 else None)

        def fmt_lr(ft_val):
            if abs(ft_val) < 1e-5:
                return ''
            inches = abs(ft_val) * 12.0
            inc16 = round(inches * 16.0) / 16.0
            whole = int(math.floor(inc16 + 1e-6))
            frac = inc16 - whole

            def frac_str(x):
                denom = 16
                num = int(round(x * denom))
                if num == 0:
                    return ''
                g = _gcd(num, denom)
                return "{}/{}".format(num//g, denom//g)
            if whole > 0 and frac > 1e-6:
                mag = "{}-{}\"".format(whole, frac_str(frac))
            elif whole > 0:
                mag = "{}\"".format(whole)
            else:
                mag = "{}\"".format(frac_str(frac))
            return (u"→" + mag) if ft_val > 0 else (u"←" + mag)

        side_arrow = fmt_lr(lateral)

        return {
            'vertical_class': vertical_class,
            'vertical_shift_ft': v_shift,
            'vertical_arrow': v_arrow,
            'side_class': side_class,
            'side_dir': side_dir,
            'side_shift_ft': lateral,
            'side_arrow': side_arrow,
            'in_centroid': ci,
            'out_centroid': co,
            'debug': {
                'zmin_in': zmin_i, 'zmax_in': zmax_i,
                'zmin_out': zmin_o, 'zmax_out': zmax_o,
                'top_diff_ft': top_diff,
                'bot_diff_ft': bot_diff,
            }
        }
