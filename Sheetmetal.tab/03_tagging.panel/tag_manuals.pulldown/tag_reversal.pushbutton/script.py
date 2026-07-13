# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from pyrevit import DB, revit, script
from config.tag_config import *

# Button info
# ==================================================
__title__ = "Reverse Tags"
__doc__ = """
Switches tags based on dictionary mapping (Left ↔ Right)
"""

# Variables
# ==================================================
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
active_view_id = uidoc.ActiveView.Id if uidoc and uidoc.ActiveView else None

swap_dictionary = {
    SLOT_SIZE_LEFT: 'Right',
    SLOT_BOD_LEFT: 'Right',
    SLOT_LENGTH_LEFT: 'Right',
    SLOT_SIZE_RIGHT: 'Left',
    SLOT_BOD_RIGHT: 'Left',
    SLOT_LENGTH_RIGHT: 'Left',
}


def safe_int_id(element_id):
    if element_id is None:
        return None

    value = getattr(element_id, 'Value', None)
    if value is not None:
        try:
            return int(value)
        except Exception:
            pass

    integer_value = getattr(element_id, 'IntegerValue', None)
    if integer_value is not None:
        try:
            return int(integer_value)
        except Exception:
            pass

    try:
        return int(element_id)
    except Exception:
        return None


def get_tagged_local_element_ids(tag):
    """Return local tagged element ids for a tag across Revit versions."""
    ids = []
    try:
        local_ids = list(tag.GetTaggedLocalElementIds() or [])
        ids.extend(local_ids)
    except Exception:
        pass

    if ids:
        return ids

    try:
        tagged_id = tag.TaggedLocalElementId
        if tagged_id and safe_int_id(tagged_id) not in (None, -1):
            ids.append(tagged_id)
    except Exception:
        pass

    return ids


selected_ids = list(uidoc.Selection.GetElementIds())
if not selected_ids:
    script.exit()

if not active_view_id:
    script.exit()

annotation_ids = set()
selected_int_ids = set(safe_int_id(eid) for eid in selected_ids)
selected_int_ids.discard(None)

view_tags = list(
    DB.FilteredElementCollector(doc, active_view_id)
    .OfClass(DB.IndependentTag)
    .ToElements()
)
for tag in view_tags:
    tagged_ids = get_tagged_local_element_ids(tag)
    for tagged_id in tagged_ids:
        if safe_int_id(tagged_id) in selected_int_ids:
            annotation_ids.add(tag.Id)
            break

for selected_id in selected_ids:
    selected_elem = doc.GetElement(selected_id)
    if isinstance(selected_elem, DB.IndependentTag):
        if active_view_id and hasattr(selected_elem, "OwnerViewId") and selected_elem.OwnerViewId != active_view_id:
            continue
        annotation_ids.add(selected_id)

if not annotation_ids:
    script.exit()


def get_param_string(element, param_name):
    """Safely get a string value from a named parameter."""
    try:
        param = element.LookupParameter(param_name)
        if not param:
            return None
        value = param.AsString()
        return value if value else param.AsValueString()
    except Exception:
        return None


def get_bip_string(element, bip_name):
    """Safely get a string value from a built-in parameter."""
    try:
        bip = getattr(DB.BuiltInParameter, bip_name, None)
        if bip is None:
            return None
        param = element.get_Parameter(bip)
        if not param:
            return None
        value = param.AsString()
        return value if value else param.AsValueString()
    except Exception:
        return None


def get_type_identity(elem_type):
    """Return (family_name, type_name) tuple for an ElementType."""
    family_name = (
        get_param_string(elem_type, "Family Name")
        or get_param_string(elem_type, "Family")
        or get_bip_string(elem_type, "SYMBOL_FAMILY_NAME_PARAM")
        or get_bip_string(elem_type, "ALL_MODEL_FAMILY_NAME")
        or getattr(elem_type, "FamilyName", None)
    )
    direct_name = None
    try:
        direct_name = DB.Element.Name.GetValue(elem_type)
    except Exception:
        direct_name = None
    type_name = (
        get_param_string(elem_type, "Type Name")
        or get_bip_string(elem_type, "SYMBOL_NAME_PARAM")
        or get_bip_string(elem_type, "ALL_MODEL_TYPE_NAME")
        or get_param_string(elem_type, "Type")
        or direct_name
        or getattr(elem_type, "Name", None)
    )
    if family_name and type_name:
        return family_name, type_name
    return None, None


def norm_text(value):
    return (value or "").strip().lower()


def get_type_pool(elem_type):
    """Return normalized (family, type, combined) strings for type matching."""
    family_name, type_name = get_type_identity(elem_type)
    family_norm = norm_text(family_name)
    type_norm = norm_text(type_name)
    pool_norm = (family_norm + " " + type_norm).strip()
    return family_norm, type_norm, pool_norm


# Build lookup from configured slot candidates:
# (family_lower, source_type_lower) -> target_type_name
configured_swaps = {}
for source_slot, target_type_name in swap_dictionary.items():
    for candidate in DEFAULT_TAG_SLOT_CANDIDATES.get(source_slot, []):
        if not isinstance(candidate, tuple) or len(candidate) < 2:
            continue
        source_family_name, source_type_name = candidate[0], candidate[1]
        configured_swaps[(norm_text(source_family_name), norm_text(source_type_name))] = target_type_name


# Build lookup: (family_lower, type_lower) -> ElementTypeId
tuple_to_type_id = {}
all_element_types = DB.FilteredElementCollector(doc).OfClass(DB.ElementType)
for elem_type in all_element_types.ToElements():
    family_name, type_name = get_type_identity(elem_type)
    key = (norm_text(family_name), norm_text(type_name))
    if key[0] and key[1]:
        tuple_to_type_id[key] = elem_type.Id


def resolve_target_type_id(tag, family_name, target_type_name):
    """Resolve a target tag type id, preferring exact family/type matches."""
    family_key = norm_text(family_name)
    target_key = norm_text(target_type_name)

    exact_id = tuple_to_type_id.get((family_key, target_key))
    if exact_id:
        return exact_id

    # Fallback: resolve within valid types for this specific tag instance.
    try:
        valid_type_ids = list(tag.GetValidTypes() or [])
    except Exception:
        valid_type_ids = []

    for type_id in valid_type_ids:
        elem_type = doc.GetElement(type_id)
        fam_norm, type_norm, _ = get_type_pool(elem_type)
        if fam_norm == family_key and type_norm == target_key:
            return type_id

    for type_id in valid_type_ids:
        elem_type = doc.GetElement(type_id)
        fam_norm, type_norm, pool_norm = get_type_pool(elem_type)
        if fam_norm == family_key and target_key in type_norm:
            return type_id
        if fam_norm == family_key and target_key in pool_norm:
            return type_id

    for (fam_norm, type_norm), type_id in tuple_to_type_id.items():
        if fam_norm == family_key and type_norm == target_key:
            return type_id

    for (fam_norm, type_norm), type_id in tuple_to_type_id.items():
        if fam_norm == family_key and target_key in type_norm:
            return type_id

    return None


def find_target_type_name(family_name, type_name):
    family_key = norm_text(family_name)
    type_key = norm_text(type_name)

    exact = configured_swaps.get((family_key, type_key))
    if exact:
        return exact

    for (src_family, src_type), target in configured_swaps.items():
        if src_family == family_key and src_type in type_key:
            return target

    return None


swapped_count = 0
skipped_count = 0
error_count = 0
with revit.Transaction("Switch tags based on dictionary"):
    for ann_id in annotation_ids:
        ann = doc.GetElement(ann_id)
        if not isinstance(ann, DB.IndependentTag):
            continue

        current_type = doc.GetElement(ann.GetTypeId())
        family_name, type_name = get_type_identity(current_type)
        target_type_name = find_target_type_name(family_name, type_name)

        if not target_type_name:
            skipped_count += 1
            continue

        target_tag_id = resolve_target_type_id(ann, family_name, target_type_name)
        if not target_tag_id:
            skipped_count += 1
            continue

        if target_tag_id == ann.GetTypeId():
            skipped_count += 1
            continue

        try:
            ann.ChangeTypeId(target_tag_id)
            swapped_count += 1
        except Exception:
            error_count += 1
