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
    ElementId,
    FilteredElementCollector,
    FamilySymbol,
    IndependentTag,
    Transaction,
    XYZ,
)
from tagging.revit_tagging import RevitTagging
from tagging.tag_config import DEFAULT_TAG_SLOT_CANDIDATES, SLOT_MARK, SLOT_MARK_NOTE

# Button display information
# =================================================
__title__ = "Tag Equipment"
__doc__ = """
Tags all mechanical equipment in the current view.
"""

# Helpers
# ==================================================
tags_mark = DEFAULT_TAG_SLOT_CANDIDATES[SLOT_MARK]
tags_mark_note = DEFAULT_TAG_SLOT_CANDIDATES[SLOT_MARK_NOTE]

MARK = SLOT_MARK
MARK_NOTE = SLOT_MARK_NOTE
TRIGGER_PARAM = '_note'


# Tag categories compatible with OST_MechanicalEquipment
_EQUIPMENT_TAG_CATEGORIES = [
    BuiltInCategory.OST_MechanicalEquipmentTags,
    BuiltInCategory.OST_MultiCategoryTags,
]


def _find_tag_symbol(doc, target_name):
    """Return the first tag symbol whose name contains target_name."""
    if not target_name:
        return None
    needle = target_name.strip().lower()
    categories = _EQUIPMENT_TAG_CATEGORIES

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


# Code
# ==================================================
uidoc = __revit__.ActiveUIDocument

doc = revit.doc
view = revit.active_view
output = script.get_output()

tagger = RevitTagging(doc, view)


def _find_first_available_tag(doc, tag_names):
    """Try to find the first available tag from a list of tag names.
    Falls back to any loaded equipment or multi-category tag."""
    for tag_name in tag_names:
        tag_sym = _find_tag_symbol(doc, tag_name)
        if tag_sym:
            return tag_sym, tag_name
    # Fallback: use any available equipment-compatible tag
    for bic in _EQUIPMENT_TAG_CATEGORIES:
        syms = list(
            FilteredElementCollector(doc)
            .OfCategory(bic)
            .OfClass(FamilySymbol)
            .ToElements()
        )
        if syms:
            fam = getattr(syms[0], "Family", None)
            fallback_name = fam.Name if fam else getattr(syms[0], "Name", "")
            return syms[0], fallback_name
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

selected_mark_symbol, selected_mark_name = _find_first_available_tag(doc, tags_mark)
if not selected_mark_symbol:
    output.print_md(
        "## No MARK tag found from configured tags list: {}".format(", ".join(tags_mark)))
    script.exit()

selected_mark_note_symbol = None
selected_mark_note_name = None
for tag_name in tags_mark_note:
    tag_sym = _find_tag_symbol(doc, tag_name)
    if tag_sym:
        selected_mark_note_symbol = tag_sym
        selected_mark_note_name = tag_name
        break

placed = []
failed = []
already_tagged = []
placed_mark = []
placed_mark_note = []
removed_conflicting = []

# Check how many tags already exist in the view
existing_tags = list(
    FilteredElementCollector(doc, view.Id)
    .OfClass(IndependentTag)
    .ToElements()
)

# Build maps of element IDs keyed by tracked tag type id.
existing_tag_maps = {}

tracked_tag_type_ids = {selected_mark_symbol.Id.IntegerValue}
if selected_mark_note_symbol:
    tracked_tag_type_ids.add(selected_mark_note_symbol.Id.IntegerValue)

mark_type_id_val = selected_mark_symbol.Id.IntegerValue
mark_note_type_id_val = (
    selected_mark_note_symbol.Id.IntegerValue if selected_mark_note_symbol else None
)


def _eid_int(eid):
    try:
        return eid.IntegerValue
    except Exception:
        try:
            return int(eid)
        except Exception:
            return None


def _collect_tagged_local_ids(tag):
    """Return local element IDs tagged by an IndependentTag (version-safe)."""
    ids = []

    # Revit 2026+ primary API
    try:
        for tid in tag.GetTaggedLocalElementIds() or []:
            if tid and tid != ElementId.InvalidElementId:
                ids.append(tid)
    except Exception:
        pass

    # Revit 2022-2025 property API
    try:
        tid = tag.TaggedLocalElementId
        if tid and tid != ElementId.InvalidElementId:
            ids.append(tid)
    except Exception:
        pass

    # Some tag types expose LinkElementId-based APIs
    def _append_from_link_eid(link_eid):
        if not link_eid:
            return
        for attr in ("HostElementId", "LinkedElementId", "ElementId"):
            try:
                candidate = getattr(link_eid, attr)
                if candidate and candidate != ElementId.InvalidElementId:
                    ids.append(candidate)
            except Exception:
                pass

    try:
        for leid in tag.GetTaggedElementIds() or []:
            _append_from_link_eid(leid)
    except Exception:
        pass

    try:
        _append_from_link_eid(tag.TaggedElementId)
    except Exception:
        pass

    # Deduplicate by int value
    uniq = []
    seen = set()
    for eid in ids:
        ival = _eid_int(eid)
        if ival is None or ival in seen:
            continue
        seen.add(ival)
        uniq.append(eid)
    return uniq


for tag in existing_tags:
    try:
        # Only count tags whose type matches selected MARK / MARK_COMMENT symbols.
        try:
            tag_type_id = tag.GetTypeId()
        except Exception:
            tag_type_id = None
        tag_type_id_val = _eid_int(tag_type_id)
        if tag_type_id is None or tag_type_id_val not in tracked_tag_type_ids:
            continue

        # Use integer -1 check (IronPython-safe) instead of ElementId object comparison.
        elem_id_int = None
        try:
            v = _eid_int(tag.TaggedLocalElementId)
            if v is not None and v != -1:
                elem_id_int = v
        except Exception:
            pass
        if elem_id_int is None:
            try:
                for tid in (tag.GetTaggedLocalElementIds() or []):
                    v = _eid_int(tid)
                    if v is not None and v != -1:
                        elem_id_int = v
                        break
            except Exception:
                pass
        if elem_id_int is not None:
            existing_tag_maps.setdefault(tag_type_id_val, {}).setdefault(elem_id_int, []).append(tag)
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

        comments_value = ""
        try:
            comments_param = elem.LookupParameter(TRIGGER_PARAM)
            if comments_param:
                comments_value = (comments_param.AsString() or comments_param.AsValueString() or "").strip()
        except Exception:
            comments_value = ""

        if comments_value and selected_mark_note_symbol:
            tag_symbol = selected_mark_note_symbol
            tag_name = selected_mark_note_name
        else:
            tag_symbol = selected_mark_symbol
            tag_name = selected_mark_name

        # Skip if element already has the chosen tag type in this view.
        elem_id_val = elem.Id.IntegerValue
        chosen_type_id_val = _eid_int(tag_symbol.Id)
        existing_for_elem = existing_tag_maps.get(chosen_type_id_val, {}).get(elem_id_val, [])
        if existing_for_elem:
            already_tagged.append(elem)
            continue

        opposite_type_id_val = None
        if chosen_type_id_val == mark_type_id_val and mark_note_type_id_val is not None:
            opposite_type_id_val = mark_note_type_id_val
        elif chosen_type_id_val == mark_note_type_id_val:
            opposite_type_id_val = mark_type_id_val

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
            new_tag = tagger.place_tag(elem, tag_symbol, tag_pt)
            placed.append(elem)
            existing_tag_maps.setdefault(chosen_type_id_val, {}).setdefault(elem_id_val, []).append(new_tag)

            # If we switched tag types, remove the old opposite tag(s) on this element.
            if opposite_type_id_val is not None:
                old_tags = list(
                    existing_tag_maps.get(opposite_type_id_val, {}).get(elem_id_val, [])
                )
                # Fallback: direct scan in case pre-built map missed any.
                if not old_tags:
                    for et in existing_tags:
                        try:
                            if _eid_int(et.GetTypeId()) != opposite_type_id_val:
                                continue
                            et_elem = None
                            try:
                                v = _eid_int(et.TaggedLocalElementId)
                                if v is not None and v != -1:
                                    et_elem = v
                            except Exception:
                                pass
                            if et_elem is None:
                                try:
                                    for tid in (et.GetTaggedLocalElementIds() or []):
                                        v = _eid_int(tid)
                                        if v is not None and v != -1:
                                            et_elem = v
                                            break
                                except Exception:
                                    pass
                            if et_elem == elem_id_val:
                                old_tags.append(et)
                        except Exception:
                            pass
                for old_tag in old_tags:
                    try:
                        doc.Delete(old_tag.Id)
                        removed_conflicting.append(old_tag)
                    except Exception:
                        pass
                existing_tag_maps.setdefault(opposite_type_id_val, {})[elem_id_val] = []

            if tag_symbol.Id == selected_mark_symbol.Id:
                placed_mark.append(elem)
            elif selected_mark_note_symbol and tag_symbol.Id == selected_mark_note_symbol.Id:
                placed_mark_note.append(elem)
        except Exception as e:
            failed.append((elem, "Tag placement error [{}]: {}".format(tag_name, str(e))))

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

output.print_md(
    "## Placed by type: MARK {}, MARK_NOTE {}".format(
        len(placed_mark),
        len(placed_mark_note),
    )
)

output.print_md(
    "## Replaced old opposite tags: {}".format(
        len(removed_conflicting),
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
