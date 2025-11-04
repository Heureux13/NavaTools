# script.py
# Copyright (c) 2025 Jose Francisco Nava Perez
# All rights reserved. No part of this code may be reproduced without permission.

# This is intended only for straight fabrication duct
from System.Collections.Generic import List
from pyrevit import revit, DB
from Autodesk.Revit.DB import FabricationPart

# Setup our basic variables
doc = revit.doc
view = revit.active_view

# Collect ducts in view
ducts = (DB.FilteredElementCollector(doc, view.Id)
         .OfCategory(DB.BuiltInCategory.OST_FabricationDuctwork)
         .WhereElementIsNotElementType()
         .ToElements())

# Collect labels in model
tag_syms = (DB.FilteredElementCollector(doc)
            .OfClass(DB.FamilySymbol)
            .OfCategory(DB.BuiltInCategory.OST_FabricationDuctworkTags)
            .ToElements())

# --- Helpers ---------------------------------------------------------------


def get_label(name_contains):
    tag = name_contains.lower()
    for ts in tag_syms:
        fam = getattr(ts, "Family", None)
        fam_name = fam.Name if fam else ""
        ts_name = getattr(ts, "Name", "")
        pool = (fam_name + " " + ts_name).lower()
        if tag in pool:
            return ts
    print("No label found with:", name_contains)
    for ts in tag_syms:
        fam = getattr(ts, "Family", None)
        fam_name = fam.Name if fam else "?"
        ts_name = getattr(ts, "Name", "?")
        print("Available tag symbol:", fam_name, "/", ts_name)
    return None


# Grab the length tag
len_tag = get_label("length")
if not len_tag:
    raise Exception(
        "Cannot find length label, check spelling or load it into the project.")


def place(element, tag_symbol, point_xyz):
    ref = DB.Reference(element)
    tag = DB.IndependentTag.Create(
        doc,
        view.Id,
        ref,
        False,  # leader off
        DB.TagMode.TM_ADDBY_CATEGORY,
        DB.TagOrientation.Horizontal,
        point_xyz
    )
    if tag_symbol and tag_symbol.Id:
        tag.ChangeTypeId(tag_symbol.Id)
    return tag


def project_to_plane(v, n):
    dot = v.DotProduct(n)
    return DB.XYZ(v.X - dot * n.X, v.Y - dot * n.Y, v.Z - dot * n.Z)


def signed_angle_in_view(direction, view):
    r = view.RightDirection.Normalize()
    n = view.ViewDirection.Normalize()
    d = direction.Normalize()
    d_plane = project_to_plane(d, n)
    if d_plane.GetLength() < 1e-9:
        return None
    d_unit = d_plane.Normalize()
    angle = d_unit.AngleTo(r)
    cross = r.CrossProduct(d_unit)
    if cross.DotProduct(n) < 0:
        angle = -angle
    return angle


def view_plane_bottom_left(duct, view):
    bbox = duct.get_BoundingBox(view) or duct.get_BoundingBox(None)
    if not bbox:
        return None, None
    r = view.RightDirection.Normalize()
    u = view.UpDirection.Normalize()
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
    def proj(p, a): return p.X * a.X + p.Y * a.Y + p.Z * a.Z
    best, best_up, best_right = None, None, None
    for c in corners:
        pu = proj(c, u)
        pr = proj(c, r)
        if best is None or pu < best_up - 1e-9 or (abs(pu - best_up) <= 1e-9 and pr < best_right):
            best, best_up, best_right = c, pu, pr
    centroid = (bbox.Min + bbox.Max) / 2.0
    return best, centroid


def rotate_tag_to_curve(tag, curve, view):
    r = view.RightDirection.Normalize()
    n = view.ViewDirection.Normalize()
    direction = curve.Direction.Normalize()
    dot = direction.DotProduct(n)
    d_plane = DB.XYZ(direction.X - dot * n.X,
                     direction.Y - dot * n.Y,
                     direction.Z - dot * n.Z)
    if d_plane.GetLength() < 1e-9:
        angle = 0.0
    else:
        d_unit = d_plane.Normalize()
        angle = d_unit.AngleTo(r)
        cross = r.CrossProduct(d_unit)
        if cross.DotProduct(n) < 0:
            angle = -angle
    loc = tag.Location
    if isinstance(loc, DB.LocationPoint):
        axis = DB.Line.CreateBound(loc.Point, loc.Point + n)
        loc.Rotate(axis, angle)


def already_tagged(elem, tag_fam_name):
    existing = (DB.FilteredElementCollector(doc, view.Id)
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


# --- CONFIG ---
OFFSET_IN = 0.15
EPS = 1e-6

# --- Main ------------------------------------------------------------------

t = DB.Transaction(doc, "Tag All Straight Duct")
t.Start()

if not len_tag.IsActive:
    len_tag.Activate()
    doc.Regenerate()

for duct in ducts:
    print("\n--- Checking duct:", duct.Id.Value, "---")
    if not isinstance(duct, FabricationPart):
        print("  Skipped: not a FabricationPart")
        continue
    if already_tagged(duct, len_tag.Family.Name):
        print("  Skipped: already tagged")
        continue
    loc = getattr(duct, "Location", None)
    if not isinstance(loc, DB.LocationCurve):
        print("  Skipped: no LocationCurve")
        continue

    curve = loc.Curve
    corner, centroid = view_plane_bottom_left(duct, view)
    if corner is None:
        print("  Skipped: no bbox")
        continue

    vec_in = centroid - corner
    placement = corner if vec_in.GetLength() < EPS else corner + \
        vec_in.Normalize().Multiply(OFFSET_IN)

    print("  corner (view bottom-left):",
          (round(corner.X, 3), round(corner.Y, 3), round(corner.Z, 3)))
    print("  centroid:", (round(centroid.X, 3),
          round(centroid.Y, 3), round(centroid.Z, 3)))
    print("  placement:", (round(placement.X, 3),
          round(placement.Y, 3), round(placement.Z, 3)))

    tag = place(duct, len_tag, placement)

    # Force the head position to the placement point (leaderless tags respect this)
    try:
        tag.TagHeadPosition = placement
    except Exception as e:
        print("  TagHeadPosition set failed:", e)

    doc.Regenerate()
    rotate_tag_to_curve(tag, curve, view)

    # Diagnostics
    loc = tag.Location
    if isinstance(loc, DB.LocationPoint):
        print("  tag head:", (round(tag.TagHeadPosition.X, 3),
                              round(tag.TagHeadPosition.Y, 3),
                              round(tag.TagHeadPosition.Z, 3)))

print("Ducts found:", len(ducts))
print("Tag symbols found:", len(tag_syms))

t.Commit()
