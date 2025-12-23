# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from System.Collections.Generic import List
from revit_output import print_disclaimer
from revit_tagging import RevitTagging
from revit_element import RevitElement
from revit_duct import RevitDuct
from pyrevit import DB, forms, revit, script
from Autodesk.Revit.DB import ElementId, Transaction

# Button info
# ==================================================
__title__ = "Fittings"
__doc__ = """
Tag all fitting with assosiated label
"""

# Variables
# ==================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
output = script.get_output()
view = revit.active_view
tagger = RevitTagging(doc=doc, view=view)

elbow_throat_allowances = {
    'tdf': 6,
    's&d': 4,
}

elbow_tag_excluted = {
    '-FabDuct_EXT IN_MV_Tag',
    '-FabDuct_EXT OUT_MV_Tag',
    '-FabDuct_EXT LEFT_MV_Tag',
    '-FabDuct_EXT RIGHT_MV_Tag'
}

elbow_families = {
    'elbow',
    'tee',
}

all_connector_types = {
    'duct.connector_0_type',
    'duct.connector_1_type',
    'duct.connector_2_type',
    'duct.connector_3_type',
}

family_to_angle_skip = {
    'radius elbow',
    'gored elbow'
}


def should_skip_tag(duct, tag):
    fam = (duct.family or '').strip().lower()
    tag_name = (tag.Family.Name if tag and tag.Family else "").strip().lower()
    # Skip -FabDuct_DEGREE_MV_Tag for Radius Elbow with angle 45 or 90
    if fam in family_to_angle_skip and duct.angle in [45, 90] and tag_name == '-fabduct_degree_mv_tag':
        return True
    # Skip extension tags for elbows/tees with tdf connector and extension == 6
    if fam in elbow_families:
        for connector_types in [duct.connector_0_type, duct.connector_1_type]:
            if not connector_types:
                continue
            connector_type_keys = connector_types.lower().strip()
            required_ext = elbow_throat_allowances.get(connector_type_keys)
            if (
                required_ext is not None
                and (
                    duct.extension_top == required_ext
                    or duct.extension_bottom == required_ext
                )
                and tag_name in {t.strip().lower() for t in elbow_tag_excluted}
            ):
                return True
    return False


# Collect ducts in view
# ==================================================
ducts = RevitDuct.all(doc, view)

# for d in ducts:
#     output.print_md("ID: {} | Fa: {} | An: {} | Ex: {}".format(d.element.Id, d.family, d.angle, d.extension_bottom))

if not ducts:
    output.print_md("No ducts found in the current view", exitscript=True)

# Dictionary: Family name: list of (tag, location) tuples
# ==================================================
duct_families = {
    "8inch long coupler wdamper": [
        (tagger.get_label("-FabDuct_TM_MV_Tag"), 0.5)
    ],

    "conical tap - wdamper": [
        (tagger.get_label("-FabDuct_TM_MV_Tag"), 0.5)
    ],

    # Rectangle tap usually on the main trunk.
    "boot tap": [
        (tagger.get_label("-FabDuct_TM_MV_Tag"), 0.5)
    ],

    # Round tap usually from main to VAV.
    "boot tap - wdamper": [
        (tagger.get_label("-FabDuct_TM_MV_Tag"), 0.5)
    ],

    # Round tap usually from main to VAV.
    "boot saddle tap": [
        (tagger.get_label("-FabDuct_TM_MV_Tag"), 0.5)
    ],

    "cap": [
        (tagger.get_label("-FabDuct_TM_MV_Tag"), 0.5)
    ],

    # Offset Radius elbow
    'drop check': [
        (tagger.get_label('-FabDuct_SIZE_FIX_Tag'), 0.5)
    ],

    # Square elbows from 5° to 90+°
    "elbow": [
        (tagger.get_label("-FabDuct_EXT IN_MV_Tag"), 0.5),
        (tagger.get_label("-FabDuct_EXT OUT_MV_Tag"), 0.5),
        (tagger.get_label("-FabDuct_DEGREE_MV_Tag"), 0.5),
    ],

    # Square elbows from 5° to 90+°
    "elbow 90 degree": [
        (tagger.get_label("-FabDuct_EXT IN_MV_Tag"), 0.5),
        (tagger.get_label("-FabDuct_EXT OUT_MV_Tag"), 0.5),
        (tagger.get_label("-FabDuct_DEGREE_MV_Tag"), 0.5),
    ],

    # Round/square/rectangle end cap
    "end cap": [
        (tagger.get_label("-FabDuct_TM_MV_Tag"), 0.5)
    ],

    # 90° adjustable elbow
    "gored elbow": [
        (tagger.get_label("-FabDuct_DEGREE_MV_Tag"), 0.5)
    ],

    "mitred offset": [
        (tagger.get_label("-FabDuct_TRAN_MV_Tag"), 0.5)
    ],

    "cid330 - (radius 2-way offset)": [
        (tagger.get_label("-FabDuct_TRAN_MV_Tag"), 0.5)
    ],

    # Square/rectangle to square/rectangle
    "offset": [
        (tagger.get_label("-FabDuct_TRAN_MV_Tag"), 0.5)
    ],

    # Offset ogee
    "ogee": [
        (tagger.get_label("-FabDuct_TRAN_MV_Tag"), 0.5)
    ],

    "radius bend": [
        (tagger.get_label("-FabDuct_SIZE_FIX_Tag"), 0.5)
    ],

    # Elbow with radius heel and throat
    "radius elbow": [
        (tagger.get_label("-FabDuct_INNER_R_FIX_Tag"), 0.5),
        (tagger.get_label('-FabDuct_DEGREE_MV_Tag'), 0.5),
    ],

    "radius offset": [
        (tagger.get_label("-FabDuct_TRAN_MV_Tag"), 0.5)
    ],

    # Round reducer
    "reducer": [
        (tagger.get_label("-FabDuct_TRAN_MV_Tag"), 0.5)
    ],

    "square bend": [
        (tagger.get_label("-FabDuct_SIZE_FIX_Tag"), 0.5)
    ],

    # Square to round
    "square to ø": [
        (tagger.get_label("-FabDuct_TRAN_MV_Tag"), 0.5)
    ],

    "tap": [
        (tagger.get_label("-FabDuct_SIZE_FIX_Tag"), 0.5)
    ],

    # TDF end cap
    "tdf end cap": [
        (tagger.get_label("-FabDuct_SIZE_FIX_Tag"), 0.5)
    ],

    # Square/rectangle tee elbow
    "tee": [
        (tagger.get_label("-FabDuct_EXT IN_MV_Tag"), 0.5),
        (tagger.get_label("-FabDuct_EXT LEFT_MV_Tag"), 0.5),
        (tagger.get_label("-FabDuct_EXT RIGHT_MV_Tag"), 0.5)
    ],

    # Square/retangele to square/rectangle reducer
    "transition": [
        (tagger.get_label("-FabDuct_TRAN_MV_Tag"), 0.5)
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

        tagged_this_element = False
        # Track existing tag families on this element (case-insensitive) to avoid duplicates
        existing_tag_fams = tagger.get_existing_tag_families(d.element)

        for tag, dic_duct_loc in tag_configs:
            if should_skip_tag(d, tag):
                continue
            fam_name = (tag.Family.Name if tag and tag.Family else "").strip().lower()
            if not fam_name:
                continue

            # Skip if a tag with this family name is already on the element (in this view)
            if fam_name in existing_tag_fams:
                continue

            # Tag placement logic
            if isinstance(d.element, DB.FabricationPart):
                face_ref, face_pt = tagger.get_face_facing_view(
                    d.element, prefer_point=None)
                if face_ref is not None and face_pt is not None:
                    tagger.place_tag(face_ref, tag, face_pt)
                    existing_tag_fams.add(fam_name)
                    tagged_this_element = True
                    continue
                bbox = d.element.get_BoundingBox(view)
                if bbox is not None:
                    center = (bbox.Min + bbox.Max) / 2.0
                    tagger.place_tag(d.element, tag, center)
                    existing_tag_fams.add(fam_name)
                    tagged_this_element = True
                    continue
                continue
            else:
                loc = getattr(d.element, "Location", None)
                if not loc:
                    bbox = d.element.get_BoundingBox(view)
                    if bbox is not None:
                        center = (bbox.Min + bbox.Max) / 2.0
                        tagger.place_tag(d.element, tag, center)
                        existing_tag_fams.add(fam_name)
                        tagged_this_element = True
                        continue
                    continue
                if hasattr(loc, "Point") and loc.Point is not None:
                    tagger.place_tag(d.element, tag, loc.Point)
                    existing_tag_fams.add(fam_name)
                    tagged_this_element = True
                elif hasattr(loc, "Curve") and loc.Curve is not None:
                    midpoint = loc.Curve.Evaluate(dic_duct_loc, True)
                    tagger.place_tag(d.element, tag, midpoint)
                    existing_tag_fams.add(fam_name)
                    tagged_this_element = True
                else:
                    continue

        # Add to appropriate list (only once per element)
        if tagged_this_element:
            needs_tagging.append(d)
        else:
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
                "### No.{} | ID: {} | Fam: {} | Size: {} | Le: {:06.2f} | Ex: {}".format(
                    i,
                    output.linkify(d.element.Id),
                    d.family,
                    d.size,
                    d.length,
                    d.extension_bottom
                )
            )
        output.print_md("---")

    if already_tagged:
        output.print_md("## Already Tagged")
        for i, d in enumerate(already_tagged, start=1):
            output.print_md(
                "### Index {} | Size: {} | Family: {} | Length: {:06.2f} | Element ID: {}".format(
                    i,
                    d.size,
                    d.family,
                    d.length,
                    output.linkify(d.element.Id)
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

    print_disclaimer(output)

    t.Commit()
except Exception as e:
    output.print_md("Tag placement error: {}".format(e))
    t.RollBack()
    raise
