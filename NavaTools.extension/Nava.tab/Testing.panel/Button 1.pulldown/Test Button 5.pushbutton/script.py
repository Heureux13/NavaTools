# script.py
# Copyright (c) 2025 Jose Francisco Nava Perez
# All rights reserved. No part of this code may be reproduced without permission.

# This is inteneded only for straight duct,

from tag_location import Point, Vector, Line
import math
from System.Collections.Generic import List
import clr
from pyrevit import revit, DB

import sys
import os


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

# Gets our labesl if they exist


def get_label(name_contains):
    key = name_contains.lower()
    for ts in tag_syms:
        fam = getattr(ts, "Family", None)
        fam_name = fam.Name if fam else ""
        ts_name = getattr(ts, "Name", "")
        pool = (fam_name + " " + ts_name).lower()
        if key in pool:
            return ts
    # Debug print if nothing matched
    print("No lable found with:", name_contains)
    for ts in tag_syms:
        fam = getattr(ts, "Family", None)
        fam_name = fam.Name if fam else "?"
        ts_name = getattr(ts, "Name", "?")
        print("Available tag symbol:", fam_name, "/", ts_name)
    return None


# Grab these labels from the model
size_tag = get_label("size")
len_tag = get_label("length")
bod_tag = get_label("bod")

if not all([size_tag, len_tag, bod_tag]):
    raise Exception(
        "Cannot find label, check spelling or make sure labels are loaded into the project.")

# Gives us the ability to start making changes to the model, the sting is just what shows in the undo menu
t = DB.Transaction(doc, "Frank's Taggabalooza - Straight Duct")
t.Start()

# Activate tag symbols if needed and refreshes the model so you can use the tags.
for ts in (size_tag, len_tag, bod_tag):
    if not ts.IsActive:
        ts.Activate()
        doc.Regenerate()

# Checks to see if the duct is already tagged with the tags we want to tag with.


def already_tagged(elem, tag_fam_names):
    existing = (DB.FilteredElementCollector(doc, view.Id)
                .OfClass(DB.IndependentTag)
                .ToElements())
    famset = set(tag_fam_names)
    for itag in existing:
        try:
            ref = itag.GetTaggedLocalElement()
        except:
            ref = None
        if ref and ref.Id == elem.Id:
            famname = itag.GetType().FamilyName
            if famname in famset:
                return True
    return False


def rotate_tag_to_curve(tag, curve):
    direction = curve.Direction.Normalize()
    xaxis = DB.XYZ(1, 0, 0)
    angle = direction.AngleTo(xaxis)

    # Cross product to determine sign
    cross = xaxis.CrossProduct(direction)
    if cross.Z < 0:
        angle = -angle

    loc = tag.Location
    if isinstance(loc, DB.LocationPoint):
        axis = DB.Line.CreateBound(loc.Point, loc.Point + DB.XYZ(0, 0, 1))
        loc.Rotate(axis, angle)


def place(elem, tsymbol, point):
    ref = DB.Reference(elem)

    tag = DB.IndependentTag.Create(
        doc,
        view.Id,
        ref,
        False,
        DB.TagMode.TM_ADDBY_CATEGORY,
        DB.TagOrientation.Horizontal,  # safe default
        point
    )

    if tag:
        try:
            tag.ChangeTypeId(tsymbol.Id)
        except Exception as e:
            print("Could not change tag type for element {}: {}".format(elem.Id, e))
        tag.HasLeader = False

        # Rotate tag once, using the helper
        rotate_tag_to_curve(tag, elem.Location.Curve)
        print("Placed tag:", tsymbol.Family.Name, type(tag.Location))

    return tag


# Main loop
for duct in ducts:
    loc = duct.Location
    if not isinstance(loc, DB.LocationCurve):
        continue

    curve = loc.Curve
    direction = curve.Direction.Normalize()
    zaxis = DB.XYZ(0, 0, 1)
    side = direction.CrossProduct(zaxis).Normalize()
    up = zaxis

    p_start = curve.Evaluate(0.05, True)
    p_mid = curve.Evaluate(0.50, True)
    p_end = curve.Evaluate(0.95, True)
    offset_pt = p_mid + side.Multiply(1.0) + up.Multiply(2.0)

    if already_tagged(duct, {size_tag.Family.Name, len_tag.Family.Name, bod_tag.Family.Name}):
        continue

    place(duct, len_tag,  p_start)
    place(duct, size_tag, offset_pt)  # now offset
    place(duct, bod_tag,  p_end)

print("Ducts found:", len(ducts))
print("Tag symbols found:", len(tag_syms))

t.Commit()
