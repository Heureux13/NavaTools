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
from revit_xyz import RevitXYZ
from pyrevit import DB, forms, revit, script
from Autodesk.Revit.DB import ElementId, Transaction, FilteredElementCollector, IndependentTag

# Button info
# ==================================================
__title__ = "Tag Fittings"
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

elbow_extension_tags = {
    '-FabDuct_EXT IN_MV_Tag',
    '-FabDuct_EXT OUT_MV_Tag',
    '-FabDuct_EXT LEFT_MV_Tag',
    '-FabDuct_EXT RIGHT_MV_Tag'
}

elbow_families = {
    'elbow',
    'tee',
    'elbow 90 sr - stamped',
}

square_elbow_families = {
    'elbow',
    'elbow 90 degree',
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

skip_parameters = {
    'mark': ['skip', 'skip n/a'],
    '_duct_tag_offset': ['skip', 'skip n/a'],
    '_duct_tag': ['skip', 'skip n/a']
}

parameter_hierarchy_to_check = [
    'mark',
    'type mark',
]

RECT_DAMPER_MARK_TAG = "-FabDuct_MARK_Tag"
RECT_DAMPER_TYPE_MARK_TAG = "-FabDuct_TM_MV_Tag"
RECT_DAMPER_SWITCH_FAMILIES = {
    RECT_DAMPER_MARK_TAG.strip().lower(),
    RECT_DAMPER_TYPE_MARK_TAG.strip().lower(),
}


def _param_value_from_element_or_type(element, param_name):
    target = (param_name or '').strip().lower()
    if not target or element is None:
        return None

    elem_type = None
    try:
        elem_type = doc.GetElement(element.GetTypeId())
    except Exception:
        elem_type = None

    for owner in [element, elem_type]:
        if owner is None:
            continue
        try:
            for p in owner.Parameters:
                try:
                    dname = p.Definition.Name if p and p.Definition else None
                    if not dname or dname.strip().lower() != target:
                        continue
                    val = p.AsString()
                    if val is None:
                        val = p.AsValueString()
                    if val is None:
                        continue
                    val = str(val).strip()
                    if val:
                        return val
                except Exception:
                    pass
        except Exception:
            pass
    return None


def _rect_volume_damper_tag_choice(duct):
    mark_param = parameter_hierarchy_to_check[0] if len(parameter_hierarchy_to_check) > 0 else 'mark'
    type_mark_param = parameter_hierarchy_to_check[1] if len(parameter_hierarchy_to_check) > 1 else 'type mark'

    mark_val = _param_value_from_element_or_type(duct.element, mark_param)
    if mark_val and mark_val.strip():
        return tagger.get_label(RECT_DAMPER_MARK_TAG), RECT_DAMPER_MARK_TAG.strip().lower()

    type_mark_val = _param_value_from_element_or_type(duct.element, type_mark_param)
    if type_mark_val and type_mark_val.strip():
        return tagger.get_label(RECT_DAMPER_TYPE_MARK_TAG), RECT_DAMPER_TYPE_MARK_TAG.strip().lower()

    # Fallback per requirement: use type-mark tag anyway if both are empty
    return tagger.get_label(RECT_DAMPER_TYPE_MARK_TAG), RECT_DAMPER_TYPE_MARK_TAG.strip().lower()


def _delete_conflicting_tags_for_element(element, keep_family_name_lower, candidate_family_names_lower):
    try:
        tags_in_view = (
            FilteredElementCollector(doc, view.Id)
            .OfClass(IndependentTag)
            .ToElements()
        )
    except Exception:
        return

    for t in tags_in_view:
        try:
            tagged_ids = t.GetTaggedLocalElementIds()
        except Exception:
            tagged_ids = []

        if not tagged_ids:
            continue

        is_for_element = False
        for tid in tagged_ids:
            try:
                tid_val = tid.Value if hasattr(tid, 'Value') else tid.IntegerValue
                if tid_val == element.Id.Value:
                    is_for_element = True
                    break
            except Exception:
                pass
        if not is_for_element:
            continue

        try:
            t_type = doc.GetElement(t.GetTypeId())
            fam_name = (t_type.Family.Name if t_type and t_type.Family else '').strip().lower()
        except Exception:
            fam_name = ''

        if fam_name in candidate_family_names_lower and fam_name != keep_family_name_lower:
            try:
                doc.Delete(t.Id)
            except Exception:
                pass


def should_skip_by_param(duct):
    for param, skip_values in skip_parameters.items():
        param_val = getattr(duct, param, None)
        if not param_val:
            param_candidates = [param, param.title(), param.upper()]
            param_val = None
            for candidate in param_candidates:
                try:
                    param_val = RevitElement(doc, view, duct.element).get_param(candidate)
                except Exception:
                    param_val = None
                if param_val is not None:
                    break
            if param_val is None:
                try:
                    type_element = doc.GetElement(duct.element.GetTypeId())
                except Exception:
                    type_element = None
                if type_element is not None:
                    for candidate in param_candidates:
                        try:
                            param_val = RevitElement(doc, view, type_element).get_param(candidate)
                        except Exception:
                            param_val = None
                        if param_val is not None:
                            break
        if param_val is None:
            continue
        param_val_str = str(param_val).strip().lower()
        if not param_val_str:
            continue
        for skip_val in skip_values:
            if param_val_str == skip_val.strip().lower():
                return True
    return False


def should_skip_tag(duct, tag):
    if should_skip_by_param(duct):
        return True

    fam = (duct.family or '').strip().lower()
    tag_name = (tag.Family.Name if tag and tag.Family else "").strip().lower()

    # Check if elbow is 45° or 90°
    is_45_or_90 = False
    if fam in square_elbow_families:
        try:
            ang = duct.angle
            if ang is not None:
                # Check exact values and float ranges
                if ang in [45, 90, 45.0, 90.0]:
                    is_45_or_90 = True
                else:
                    try:
                        ang_float = abs(float(ang))
                        if (44.5 <= ang_float <= 45.5) or (89.5 <= ang_float <= 90.5):
                            is_45_or_90 = True
                    except (ValueError, TypeError):
                        pass
        except Exception:
            pass

    # For 45/90 square elbows: tag if vertical, skip if horizontal (for degree tags)
    if is_45_or_90 and tag_name == '-fabduct_degree_mv_tag':
        try:
            conn_origins = RevitXYZ(duct.element).connector_origins()
            if conn_origins and len(conn_origins) >= 2:
                c0, c1 = conn_origins[0], conn_origins[1]
                dz = abs(c1.Z - c0.Z)
                # If dz > 0.01, it's vertical: tag it (do NOT skip)
                # If dz <= 0.01, it's horizontal: skip tagging
                if dz <= 0.01:
                    return True
        except Exception as e:
            output.print_md("DEBUG: Exception checking vertical/horizontal (square elbow): {}".format(str(e)))
            pass
        # If vertical, allow tagging (do NOT skip)
        return False

    # Skip extension tags for elbows with vertical movement (not purely horizontal)
    if fam in elbow_families and tag_name in {t.strip().lower() for t in elbow_extension_tags}:
        try:
            # Check connector Z difference - if any vertical movement, skip extension tag
            conn_origins = RevitXYZ(duct.element).connector_origins()
            if conn_origins and len(conn_origins) >= 2:
                c0, c1 = conn_origins[0], conn_origins[1]
                dz = abs(c1.Z - c0.Z)

                # Skip extension tags for any elbow with vertical movement (tolerance 0.01 ft for floating point)
                if dz > 0.01:
                    return True
        except Exception as e:
            output.print_md("DEBUG: Exception checking vertical: {}".format(str(e)))
            pass

    # For degree tags: tag vertical elbows, skip horizontal elbows
    if fam in family_to_angle_skip and duct.angle in [45, 90] and tag_name == '-fabduct_degree_mv_tag':
        try:
            conn_origins = RevitXYZ(duct.element).connector_origins()
            if conn_origins and len(conn_origins) >= 2:
                c0, c1 = conn_origins[0], conn_origins[1]
                dz = abs(c1.Z - c0.Z)
                # If dz > 0.01, it's vertical: tag it (do NOT skip)
                # If dz <= 0.01, it's horizontal: skip tagging
                if dz <= 0.01:
                    return True
        except Exception as e:
            output.print_md("DEBUG: Exception checking vertical/horizontal: {}".format(str(e)))
            pass
        # If vertical, allow tagging (do NOT skip)
        return False

    # Skip extension tags when extension equals throat allowance (TDF/S&D),
    # with tolerance and connector type synonyms handled.
    if fam in elbow_families and tag_name in {t.strip().lower() for t in elbow_extension_tags}:
        # Normalize connector type names and include potential synonyms
        connector_types = [duct.connector_0_type, duct.connector_1_type, getattr(duct, 'connector_2_type', None)]
        for ctype in connector_types:
            if not ctype:
                continue
            key = ctype.lower().strip()
            # Map common variants to base keys
            if key in {'slip & drive', 'standing s&d'}:
                key = 's&d'

            required_ext = elbow_throat_allowances.get(key)
            if required_ext is None:
                continue

            # Compare against all four extension params with a small tolerance
            tol = 0.01
            ext_vals = [
                duct.extension_top,
                duct.extension_bottom,
                getattr(duct, 'extension_left', None),
                getattr(duct, 'extension_right', None),
            ]
            for ev in ext_vals:
                if isinstance(ev, (int, float)) and abs(ev - required_ext) <= tol:
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
    "rect volume damper": [
        (tagger.get_label("-FabDuct_TM_MV_Tag"), 0.5)
    ],

    # Round tap usually from main to VAV.
    "boot tap - wdamper": [
        (tagger.get_label("-FabDuct_TM_MV_Tag"), 0.5)
    ],

    # Round tap usually from main to VAV.
    "access panel": [
        (tagger.get_label("-FabDuct_TM_MV_Tag"), 0.5)
    ],

    "cap": [
        (tagger.get_label("-FabDuct_TM_MV_Tag"), 0.5)
    ],

    "access panel": [
        (tagger.get_label("-FabDuct_TM_MV_Tag"), 0.5)
    ],

    # Offset Radius elbow
    'drop cheek': [
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

    # Fire damerps that are tyep b
    "fire damper - type b": [
        (tagger.get_label("-FabDuct_MARK_Tag"), 0.5)
    ],

    # Man bars, Security Bars, Burglar Bars
    "manbars": [
        (tagger.get_label("-FabDuct_MARK_Tag"), 0.5)
    ],

    # Man bars, Security Bars, Burglar Bars
    "canvas": [
        (tagger.get_label("-FabDuct_TM_MV_Tag"), 0.5)
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
        (tagger.get_label('-FabDuct_DEGREE_MV_Tag'), 0.5)
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
        (tagger.get_label("-FabDuct_TM_MV_Tag"), 0.5)
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
    skipped_by_param = []

    for d in dic_ducts:
        key = d.family.strip().lower() if d.family else None
        tag_configs = duct_families.get(key)
        if not tag_configs:
            continue

        if should_skip_by_param(d):
            skipped_by_param.append(d)
            continue

        tagged_this_element = False
        # Track existing tag families on this element (case-insensitive) to avoid duplicates
        existing_tag_fams = tagger.get_existing_tag_families(d.element)

        if key == 'rect volume damper':
            chosen_tag, chosen_family_name = _rect_volume_damper_tag_choice(d)
            if chosen_tag is not None:
                _delete_conflicting_tags_for_element(
                    d.element,
                    chosen_family_name,
                    RECT_DAMPER_SWITCH_FAMILIES
                )
                existing_tag_fams = tagger.get_existing_tag_families(d.element)
                tag_configs = [(chosen_tag, 0.5)]

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

    if skipped_by_param:
        output.print_md("## Skipped by Parameter")
        for i, d in enumerate(skipped_by_param, start=1):
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
    if skipped_by_param:
        skipped_ids = [d.element.Id for d in skipped_by_param]
        output.print_md("# Skipped by parameter: {}, {}".format(
            len(skipped_by_param), output.linkify(skipped_ids)))
    all_ids = [d.element.Id for d in dic_ducts]
    output.print_md("# Total: {}, {}".format(
        len(dic_ducts), output.linkify(all_ids)))

    print_disclaimer(output)

    t.Commit()
except Exception as e:
    output.print_md("Tag placement error: {}".format(e))
    t.RollBack()
    raise
