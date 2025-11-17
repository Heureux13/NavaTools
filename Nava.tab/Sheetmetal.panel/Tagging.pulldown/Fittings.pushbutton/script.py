# -*- coding: utf-8 -*-
# ======================================================================
# Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.
#
# This code and associated documentation files may not be copied, modified,
# distributed, or used in any form without the prior written permission of
# the copyright holder.
# ======================================================================

# Imports
# ==================================================
from Autodesk.Revit.ApplicationServices import Application
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI import UIDocument
from pyrevit import revit, forms, DB, script
from revit_element import RevitElement
from revit_duct import RevitDuct
from revit_tagging import RevitTagging, TagConfig

# Button info
# ==================================================
__title__ = "Fittings"
__doc__ = """
************************************************************************
Description:
Select all fitting that need to be tagged, and tag them.w
************************************************************************
"""

# Variables
# ==================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
output = script.get_output()
view = revit.active_view
tagger = RevitTagging(doc=doc, view=view)
mid_p = RevitTagging.midpoint_location

# Collect ducts in view
# ==================================================
ducts = RevitDuct.all(doc, view)
if not ducts:
    forms.alert("No ducts found in the current view", exitscript=True)


# --- Configuration Objects -------------------------------------------------
# Example custom location function (optional)
configs = [
    TagConfig(
        names=("radius elbow", "goard elbow"),
        tags=[
            (tagger.get_label("_umi_radius_inner"), 0.5, 0.0),
        ],
        # Example: only radius elbows not exactly 90°
        predicate=lambda d: d.angle is None or abs(d.angle - 90.0) > 0.01,
        location_func=mid_p,
    ),
    TagConfig(
        names=("drop cheek bend", "radius tee"),
        tags=[
            (tagger.get_label("_umi_radius_inner"), 0.5, 0.0),
            (tagger.get_label("_umi_angle"), 0.5, 0.0),
        ],
        # Example: require inner radius to exist
        predicate=lambda d: d.inner_radius is not None,
    ),
    TagConfig(
        names=("elbow",),
        tags=[
            (tagger.get_label("_umi_angle"), 0.5, 0.0),
        ],
        # Only elbows longer than 12"
        predicate=lambda d: d.length and d.length > 12,
    ),
    TagConfig(
        names=(
            "cap",
            "end cap",
            "conical tap - wdamper",
            "8inch long coupler wdamper",
            "boot tap - wdamper",
        ),
        tags=[(tagger.get_label("_umi_size"), 0.5, 0.0)],
    ),
    TagConfig(
        names=("tee",),
        tags=[
            (tagger.get_label("_umi_extension_bottom"), 0.5, 0.0),
            (tagger.get_label("_umi_extension_left"), 0.5, 0.0),
            (tagger.get_label("_umi_extension_right"), 0.5, 0.0),
        ],
        # Example: only tag tees that have a right extension
        predicate=lambda d: d.extension_right is not None,
    ),
    TagConfig(
        names=("reducer", "transition"),
        tags=[(tagger.get_label("_umi_reducer"), 0.5, 0.0)],
    ),
    TagConfig(
        names=("mitred offset", "offset"),
        tags=[(tagger.get_label("_umi_offset"), 0.5, 0.0)],
        # Example: only offsets with length > 10"
        predicate=lambda d: d.length and d.length > 10,
    ),
]

# --- Build target list (matches + predicate) -------------------------------
target_ducts = []
for d in ducts:
    fam = (d.family or "").strip().lower()
    if not fam:
        continue
    for cfg in configs:
        if cfg.matches(fam) and cfg.predicate(d):
            target_ducts.append((d, cfg))
            break  # stop after first matching config

if not target_ducts:
    forms.alert("No matching fittings found.", exitscript=True)

# --- Transaction & Tag Placement -------------------------------------------
t = Transaction(doc, "Fittings Tagging")
t.Start()
try:
    for d, cfg in target_ducts:
        for tag, x_loc, z_offset in cfg.tags:
            if tagger.already_tagged(d.element, tag.Family.Name):
                output.print_md(
                    "Element {} already has tag '{}'.".format(
                        d.element.Id, tag.Family.Name
                    )
                )
                continue

            # FabricationPart logic
            if isinstance(d.element, DB.FabricationPart):
                face_ref, face_pt = tagger.get_face_facing_view(
                    d.element, prefer_point=None
                )
                if face_ref and face_pt:
                    tagger.place_tag(face_ref, tag, face_pt)
                    continue
                # fallback bbox
                bbox = d.element.get_BoundingBox(view)
                if bbox:
                    center = (bbox.Min + bbox.Max) / 2.0
                    tagger.place_tag(d.element, tag, center)
                    continue
                output.print_md("No geometry/bbox for {}".format(d.element.Id))
                continue

            # Non-fabrication location
            loc_pt = None
            if cfg.location_func:
                loc_pt = cfg.location_func(d, x_loc, z_offset)
            else:
                loc = d.element.Location
                if hasattr(loc, "Point") and loc.Point:
                    loc_pt = DB.XYZ(loc.Point.X, loc.Point.Y, loc.Point.Z + z_offset)
                elif hasattr(loc, "Curve") and loc.Curve:
                    curve_pt = loc.Curve.Evaluate(x_loc, True)
                    loc_pt = DB.XYZ(curve_pt.X, curve_pt.Y, curve_pt.Z + z_offset)
                else:
                    bbox = d.element.get_BoundingBox(view)
                    if bbox:
                        center = (bbox.Min + bbox.Max) / 2.0
                        loc_pt = DB.XYZ(center.X, center.Y, center.Z + z_offset)

            if loc_pt:
                tagger.place_tag(d.element, tag, loc_pt)
            else:
                output.print_md(
                    "Failed to compute tag point for {}".format(d.element.Id)
                )

    t.Commit()
    output.print_md("Tagging committed.")
except Exception as ex:
    output.print_md("Error: {}".format(ex))
    t.RollBack()
    raise

# --- Selection & Summary ---------------------------------------------------
RevitElement.select_many(uidoc, [d for d, _ in target_ducts])
output.print_md("Tagged candidate count: {}".format(len(target_ducts)))
