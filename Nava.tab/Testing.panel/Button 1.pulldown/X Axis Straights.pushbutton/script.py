# -*- coding: utf-8 -*-
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""



from pyrevit import revit, DB
from Autodesk.Revit.DB import FabricationPart

doc = revit.doc
view = revit.active_view

# Collect ducts in view
ducts = (DB.FilteredElementCollector(doc, view.Id)
         .OfCategory(DB.BuiltInCategory.OST_FabricationDuctwork)
         .WhereElementIsNotElementType()
         .ToElements())

# Collect labels
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
    raise Exception("No label found with: " + name_contains)


len_tag = get_label("length")


def place(element, tag_symbol, point_xyz):
    ref = DB.Reference(element)
    tag = DB.IndependentTag.Create(
        doc,
        view.Id,
        ref,
        False,
        DB.TagMode.TM_ADDBY_CATEGORY,
        DB.TagOrientation.Horizontal,
        point_xyz
    )
    if tag_symbol and tag_symbol.Id:
        tag.ChangeTypeId(tag_symbol.Id)
    return tag


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

# Projection helpers


def project_relative(p, origin, view):
    v = p - origin
    return (v.DotProduct(view.RightDirection), v.DotProduct(view.UpDirection))


def reconstruct_from_relative(u, v, origin, view):
    return origin + view.RightDirection.Multiply(u) + view.UpDirection.Multiply(v)

# --- Main ------------------------------------------------------------------


t = DB.Transaction(doc, "Tag Length - Horizontal Ducts Only")
t.Start()

if not len_tag.IsActive:
    len_tag.Activate()
    doc.Regenerate()

for duct in ducts:
    if not isinstance(duct, FabricationPart):
        continue
    if already_tagged(duct, len_tag.Family.Name):
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

    u = vec.DotProduct(view.RightDirection)  # horizontal component in view
    v = vec.DotProduct(view.UpDirection)     # vertical component in view

    print("Duct", duct.Id, "u:", u, "v:", v)

    # Only keep horizontals: u dominates, v is small
    if abs(v) > abs(u) * 0.2:
        continue

    bbox = duct.get_BoundingBox(view) or duct.get_BoundingBox(None)
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
    projected = [project_relative(c, origin, view) for c in corners]
    u_vals = [uv[0] for uv in projected]
    v_vals = [uv[1] for uv in projected]
    umin, umax = min(u_vals), max(u_vals)
    vmin, vmax = min(v_vals), max(v_vals)

    # Bottom-left corner of duct bbox
    chosen_u = umin
    chosen_v = vmin

    placement = reconstruct_from_relative(chosen_u, chosen_v, origin, view)

    # Create tag at dummy point, then force head
    tag = place(duct, len_tag, origin)
    tag.TagHeadPosition = placement

    # Force consistent orientation
    tag.TagOrientation = DB.TagOrientation.Horizontal
    tag.HasLeader = False

    doc.Regenerate()

print("Ducts found:", len(ducts))
print("Tag symbols found:", len(tag_syms))

t.Commit()
