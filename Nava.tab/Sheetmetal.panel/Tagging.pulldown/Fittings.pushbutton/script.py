# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from Autodesk.Revit.DB import ElementId, Transaction
from pyrevit import DB, forms, revit, script
from revit_duct import RevitDuct
from revit_element import RevitElement
from revit_tagging import RevitTagging
from revit_output import print_parameter_help
from System.Collections.Generic import List

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


# Collect ducts in view
# ==================================================
ducts = RevitDuct.all(doc, view)
if not ducts:
    forms.alert("No ducts found in the current view", exitscript=True)

# Dictionary: Family name: list of (tag, location) tuples
# ==================================================
duct_families = {
    "radius bend": [
        (tagger.get_label("_umi_size"), 0.5)
    ],
    "elbow": [
        (tagger.get_label("_umi_size"), 0.5)
    ],
    "conical tap - wdamper": [
        (tagger.get_label("_umi_size"), 0.5)
    ],
    "boot tap - wdamper": [
        (tagger.get_label("_umi_size"), 0.5)
    ],
    "8inch long coupler wdamper": [
        (tagger.get_label("_umi_size"), 0.5)
    ],
    "cap": [
        (tagger.get_label("_umi_size"), 0.5)
    ],
    "end cap": [
        (tagger.get_label("_umi_size"), 0.5)
    ],
    "tdf end cap": [
        (tagger.get_label("_umi_size"), 0.5)
    ],
    "square bend": [
        (tagger.get_label("_umi_size"), 0.5)
    ],
    "tee": [
        (tagger.get_label("_umi_size"), 0.5)
    ],
    "transition": [
        (tagger.get_label("_umi_offset_testing"), 0.5)
    ],
    "mitred offset": [
        (tagger.get_label("_umi_offset_testing"), 0.5)
    ],
    "reducer": [
        (tagger.get_label("_umi_offset_testing"), 0.5)
    ],
    "radius offset": [
        (tagger.get_label("_umi_offset_testing"), 0.5)
    ],
    "square to ø": [
        (tagger.get_label("_umi_offset_testing"), 0.5)
    ],
    "radius elbow": [
        (tagger.get_label("_umi_radius_inner"), 0.5)
    ],
    "gored elbow": [
        (tagger.get_label("_umi_radius_inner"), 0.5)
    ],
    "offset": [
        (tagger.get_label("_umi_offset_testing"), 0.5)
    ],
    "ogee": [
        (tagger.get_label("_umi_offset_testing"), 0.5)
    ],
    "tap": [
        (tagger.get_label("_umi_size"), 0.5)
    ],
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
    # Track status for reporting/selection
    needs_tagging = []
    already_tagged = []

    for d in dic_ducts:
        key = d.family.strip().lower() if d.family else None
        tag_configs = duct_families.get(key)
        if not tag_configs:
            continue

        # Track if element was newly tagged or already had all tags
        element_newly_tagged = False
        element_already_tagged = True

        # Place each tag for this element
        for tag, dic_duct_loc in tag_configs:
            if tagger.already_tagged(d.element, tag.Family.Name):
                continue

            element_already_tagged = False
            element_newly_tagged = True

            # Check if the element is a FabricationPart
            if isinstance(d.element, DB.FabricationPart):
                face_ref, face_pt = tagger.get_face_facing_view(
                    d.element, prefer_point=None)
                if face_ref is not None and face_pt is not None:
                    tagger.place_tag(face_ref, tag, face_pt)
                    continue

                # Fallback: bbox center
                bbox = d.element.get_BoundingBox(view)
                if bbox is not None:
                    center = (bbox.Min + bbox.Max) / 2.0
                    tagger.place_tag(d.element, tag, center)
                    continue
                continue
            else:
                # Handle other elements with location
                loc = getattr(d.element, "Location", None)
                if not loc:
                    bbox = d.element.get_BoundingBox(view)
                    if bbox is not None:
                        center = (bbox.Min + bbox.Max) / 2.0
                        tagger.place_tag(d.element, tag, center)
                        continue
                    continue
                if hasattr(loc, "Point") and loc.Point is not None:
                    tagger.place_tag(d.element, tag, loc.Point)
                elif hasattr(loc, "Curve") and loc.Curve is not None:
                    midpoint = loc.Curve.Evaluate(dic_duct_loc, True)
                    tagger.place_tag(d.element, tag, midpoint)
                else:
                    continue

        # Add to appropriate list (only once per element)
        if element_newly_tagged:
            needs_tagging.append(d)
        elif element_already_tagged:
            already_tagged.append(d)

    # Selection and reporting (standardized)
    if needs_tagging:
        RevitElement.select_many(uidoc, needs_tagging)
        output.print_md("# Tagged {} new fitting(s) | {} total fittings in view".format(
            len(needs_tagging), len(dic_ducts)))
    else:
        uidoc.Selection.SetElementIds(List[ElementId]())
        output.print_md(
            "# All {} fitting(s) were already tagged".format(len(dic_ducts)))

    output.print_md("---")

    if needs_tagging:
        output.print_md("## Newly Tagged")
        for i, d in enumerate(needs_tagging, start=1):
            output.print_md(
                "### Index {} | Size: {} | Family: {} | Length: {} | Element ID: {}".format(
                    i, d.size, d.family, d.length, output.linkify(
                        d.element.Id)
                )
            )
        output.print_md("---")

    if already_tagged:
        output.print_md("## Already Tagged")
        for i, d in enumerate(already_tagged, start=1):
            output.print_md(
                "### Index {} | Size: {} | Family: {} | Length: {} | Element ID: {}".format(
                    i, d.size, d.family, d.length, output.linkify(
                        d.element.Id)
                )
            )
        output.print_md("---")

    if needs_tagging:
        newly_ids = [d.element.Id for d in needs_tagging]
        output.print_md("# Newly tagged: {}, {}".format(
            len(needs_tagging), output.linkify(newly_ids)))
    if already_tagged:
        already_ids = [d.element.Id for d in already_tagged]
        output.print_md("# Already tagged: {}, {}".format(
            len(already_tagged), output.linkify(already_ids)))
    all_ids = [d.element.Id for d in dic_ducts]
    output.print_md("# Total: {}, {}".format(
        len(dic_ducts), output.linkify(all_ids)))

    print_parameter_help(output)

    t.Commit()
except Exception as e:
    output.print_md("Tag placement error: {}".format(e))
    t.RollBack()
    raise
