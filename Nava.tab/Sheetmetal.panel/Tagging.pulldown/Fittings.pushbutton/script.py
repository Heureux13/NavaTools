# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
import clr
from Autodesk.Revit.ApplicationServices import Application
from Autodesk.Revit.DB import ElementId, Reference, Transaction
from Autodesk.Revit.UI import UIDocument
from pyrevit import DB, forms, revit, script
from revit_duct import DuctAngleAllowance, JointSize, RevitDuct
from revit_element import RevitElement
from revit_tagging import RevitTagging
from revit_xyz import RevitXYZ
from System.Collections.Generic import List
from revit_parameter import RevitParameter

# Button info
# ==================================================
__title__ = "Fittings"
__doc__ = """
************************************************************************
Description:
Select all mitered elbows not 90° and all radius elbows.
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
rp = RevitParameter(doc, app)


# Collect ducts in view
# ==================================================
ducts = RevitDuct.all(doc, view)
if not ducts:
    forms.alert("No ducts found in the current view", exitscript=True)

# Dictionary: Family name: tag name
# ==================================================
duct_families = {
    "radius bend": (tagger.get_label("0_size"), 0.5),
    "elbow": (tagger.get_label("0_size"), 0.5),
    "conical tap - wdamper": (tagger.get_label("0_size"), 0.5),
    "boot tap - wdamper": (tagger.get_label("0_size"), 0.5),
    "8inch long coupler wdamper": (tagger.get_label("0_size"), 0.5),
    "cap": (tagger.get_label("0_size"), 0.5),
    "square bend": (tagger.get_label("0_size"), 0.5),
    "tee": (tagger.get_label("0_size"), 0.5),
    "transition": (tagger.get_label("0_offset_param"), 0.5),
    "mitred offset": (tagger.get_label("0_size"), 0.5),
    "radius offset": (tagger.get_label("0_size"), 0.5),
    "tap": (tagger.get_label("0_size"), 0.5),
}

# Filter ducts
# ==================================================
# Ensure d.family is not None before calling strip()
dic_ducts = [d for d in ducts if d.family and d.family.strip().lower()
             in duct_families]

# Transaction
# ==================================================
t = Transaction(doc, "General Tagging")
t.Start()
try:
    for d in ducts:
        tag = d.get_offset_value()

        if tag is not None:
            rp.set_parameter_value(d.element, "_Offset", tag)
    for d in dic_ducts:
        tag, dic_duct_loc = duct_families.get(d.family.strip().lower())
        if not tag:
            output.print_md("No tag found for family: '{}'".format(d.family))
            continue
        if tagger.already_tagged(d.element, tag.Family.Name):
            output.print_md(
                "Element {} is already tagged.".format(d.element.Id))
            continue

        # Check if the element is a FabricationPart
        if isinstance(d.element, DB.FabricationPart):
            output.print_md(
                "Processing FabricationPart: {}".format(d.element.Id))
            # Prefer a face reference that faces the current view
            face_ref, face_pt = tagger.get_face_facing_view(
                d.element, prefer_point=None)
            if face_ref is not None and face_pt is not None:
                tagger.place_tag(face_ref, tag, face_pt)
                continue

            # Fallback: use element bounding box center in this view (model coords)
            bbox = d.element.get_BoundingBox(view)
            if bbox is not None:
                center = (bbox.Min + bbox.Max) / 2.0
                output.print_md(
                    "Placing tag at bbox center: {}".format(center))
                tagger.place_tag(d.element, tag, center)
                continue

            output.print_md(
                "No valid geometry or bbox for FabricationPart: {}".format(d.element.Id))
            continue
        else:
            # Handle other elements with location
            loc = d.element.location
            if not loc:
                output.print_md(
                    "Element {} has no location.".format(d.element.Id))
                # Use element bounding box center if available
                bbox = d.element.get_BoundingBox(view)
                if bbox is not None:
                    center = (bbox.Min + bbox.Max) / 2.0
                    output.print_md(
                        "Placing tag at bbox center: {}".format(center))
                    tagger.place_tag(d.element, tag, center)
                    continue
                continue
            if hasattr(loc, "Point") and loc.Point is not None:
                output.print_md("Placing tag at point: {}".format(loc.Point))
                tagger.place_tag(d.element, tag, loc.Point)
            elif hasattr(loc, "Curve") and loc.Curve is not None:
                midpoint = loc.Curve.Evaluate(dic_duct_loc, True)
                output.print_md(
                    "Placing tag at curve midpoint: {}".format(midpoint))
                tagger.place_tag(d.element, tag, midpoint)
            else:
                output.print_md(
                    "No valid location found for element: {}".format(d.element.Id))
    t.Commit()
    output.print_md("Transaction committed successfully.")
except Exception as e:
    output.print_md("Tag placement error: {}".format(e))
    t.RollBack()
    raise

# Out put results
# ==================================================
output.print_md("## Selected {} short joint(s)".format(len(dic_ducts)))
output.print_md("---")

RevitElement.select_many(uidoc, dic_ducts)
output.print_md("Selected {} joints of duct".format(len(dic_ducts)))
