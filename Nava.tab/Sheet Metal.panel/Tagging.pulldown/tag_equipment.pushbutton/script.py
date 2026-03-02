# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from pyrevit import revit, script
from Autodesk.Revit.DB import (
    BuiltInCategory,
    FilteredElementCollector,
    FamilySymbol,
    IndependentTag,
    Transaction,
    XYZ,
)
from revit_tagging import RevitTagging

# Button display information
# =================================================
__title__ = "Tag Equipment"
__doc__ = """
Tags all mechanical equipment in the current view.
"""

# Helpers
# ==================================================
tags = [
    "-UMI Mechanical EQ Tag",
    "UMI Mechanical EQ Tag",
    "FabDuct_MARK_Tag",
]


def _find_tag_symbol(doc, target_name):
    """Return the first tag symbol whose name contains target_name."""
    if not target_name:
        return None
    needle = target_name.strip().lower()
    categories = [
        BuiltInCategory.OST_MechanicalEquipmentTags,
        BuiltInCategory.OST_MultiCategoryTags,
        BuiltInCategory.OST_FabricationDuctworkTags,
    ]

    symbols = []
    for bic in categories:
        symbols.extend(
            FilteredElementCollector(doc)
            .OfCategory(bic)
            .OfClass(FamilySymbol)
            .ToElements()
        )
    exact_matches = []
    contains_matches = []
    for sym in symbols:
        fam = getattr(sym, "Family", None)
        fam_name = fam.Name if fam else ""
        type_name = getattr(sym, "Name", "") or ""
        fam_norm = fam_name.strip().lower()
        type_norm = type_name.strip().lower()
        label = (fam_name + " " + type_name).lower()
        if needle == fam_norm or needle == type_norm:
            exact_matches.append(sym)
        elif needle in label:
            contains_matches.append(sym)

    if exact_matches:
        return exact_matches[0]
    if contains_matches:
        return contains_matches[0]
    return None


def _tag_type_matches_target(tag_type, target_name):
    if not tag_type or not target_name:
        return False
    target = target_name.strip().lower()
    fam = getattr(tag_type, 'Family', None)
    fam_name = (fam.Name if fam else "").strip().lower()
    type_name = (getattr(tag_type, 'Name', '') or '').strip().lower()
    if target == fam_name or target == type_name:
        return True
    label = (fam_name + " " + type_name).strip()
    return target in label


# Code
# ==================================================
uidoc = __revit__.ActiveUIDocument

doc = revit.doc
view = revit.active_view
output = script.get_output()

tagger = RevitTagging(doc, view)


def _find_first_available_tag(doc, tag_names):
    """Try to find the first available tag from a list of tag names."""
    for tag_name in tag_names:
        tag_sym = _find_tag_symbol(doc, tag_name)
        if tag_sym:
            return tag_sym, tag_name
    return None, None


equipment_elements = (
    FilteredElementCollector(doc, view.Id)
    .OfCategory(BuiltInCategory.OST_MechanicalEquipment)
    .WhereElementIsNotElementType()
    .ToElements()
)

if not equipment_elements:
    # output.print_md("## No mechanical equipment found in this view.")
    script.exit()

selected_tag_symbol, selected_tag_name = _find_first_available_tag(doc, tags)
if not selected_tag_symbol:
    output.print_md("## No tag found from configured tags list: {}".format(", ".join(tags)))
    script.exit()

selected_tag_family = getattr(selected_tag_symbol, "Family", None)
selected_tag_family_name = selected_tag_family.Name if selected_tag_family else ""

placed = []
failed = []
already_tagged = []

# Check how many tags already exist in the view
existing_tags = list(
    FilteredElementCollector(doc, view.Id)
    .OfClass(IndependentTag)
    .ToElements()
)

# Build a map of element IDs to existing tag instances from our families
tag_map = {}
all_tag_names = tags

for tag in existing_tags:
    try:
        tag_type_id = tag.GetTypeId()
        tag_type = doc.GetElement(tag_type_id)
        if not tag_type:
            continue

        is_our_tag = False
        for tag_name in all_tag_names:
            if _tag_type_matches_target(tag_type, tag_name):
                is_our_tag = True
                break

        if not is_our_tag:
            continue

        try:
            tagged_ids = tag.GetTaggedLocalElementIds()
        except Exception:
            tagged_ids = []

        for tid in tagged_ids:
            tid_val = tid.IntegerValue if hasattr(tid, 'IntegerValue') else int(tid)
            tag_map.setdefault(tid_val, []).append(tag)
    except BaseException:
        pass

t = Transaction(doc, "Tag Mechanical Equipment")
t.Start()
try:
    for elem in equipment_elements:
        try:
            if not elem.Category or elem.Category.Id.IntegerValue != int(BuiltInCategory.OST_MechanicalEquipment):
                failed.append((elem, "Skipped non-equipment category"))
                continue
        except Exception:
            failed.append((elem, "Unable to validate category"))
            continue

        tag_symbol = selected_tag_symbol
        tag_name = selected_tag_name

        # Prevent duplicate tags if the button is run multiple times
        if selected_tag_family_name and tagger.already_tagged(elem, selected_tag_family_name):
            already_tagged.append(elem)
            continue

        # Skip if already tagged with the correct tag; otherwise delete wrong tags
        elem_id_val = elem.Id.IntegerValue
        existing_for_elem = tag_map.get(elem_id_val, [])
        if existing_for_elem:
            has_correct = False
            for existing_tag in existing_for_elem:
                existing_type = doc.GetElement(existing_tag.GetTypeId())
                if _tag_type_matches_target(existing_type, tag_name):
                    has_correct = True
                    break
            if has_correct:
                already_tagged.append(elem)
                continue

            for existing_tag in existing_for_elem:
                try:
                    doc.Delete(existing_tag.Id)
                except Exception:
                    pass

        # Get location point for tag placement - use element location directly
        tag_pt = None
        try:
            loc = elem.Location
            if hasattr(loc, 'Point'):
                tag_pt = loc.Point
        except Exception:
            pass

        if tag_pt is None:
            # Fallback to bounding box center
            view = uidoc.ActiveView
            bbox = elem.get_BoundingBox(view) if view else None
            if bbox:
                min_pt = bbox.Min
                max_pt = bbox.Max
                tag_pt = XYZ(
                    (min_pt.X + max_pt.X) / 2.0,
                    (min_pt.Y + max_pt.Y) / 2.0,
                    (min_pt.Z + max_pt.Z) / 2.0,
                )
            else:
                failed.append((elem, "Unable to determine tag location"))
                continue

        # Place tag using element directly with its location point
        try:
            tagger.place_tag(elem, tag_symbol, tag_pt)
            placed.append(elem)
        except Exception as e:
            failed.append((elem, "Tag placement error: {}".format(str(e))))

    t.Commit()
except Exception as e:
    # output.print_md("Tag placement error: {}".format(e))
    t.RollBack()
    raise

output.print_md(
    "## Summary: placed {}, already tagged {}, failed {}".format(
        len(placed),
        len(already_tagged),
        len(failed),
    )
)

if failed:
    output.print_md("\n### Failed Elements:")
    for idx, (elem, reason) in enumerate(failed, 1):
        output.print_md(
            "- {:03} | ID: {} | Reason: {}".format(
                idx,
                output.linkify(elem.Id),
                reason,
            )
        )
