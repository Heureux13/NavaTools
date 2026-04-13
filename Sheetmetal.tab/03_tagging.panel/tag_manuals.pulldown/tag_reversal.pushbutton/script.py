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
from tagging.tag_config import *

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

dictonary = {
    SLOT_SIZE_LEFT: 'Right',
    SLOT_BOD_LEFT: 'Right',
    SLOT_LENGTH_LEFT: 'Right',
    SLOT_SIZE_RIGHT: 'Left',
    SLOT_BOD_RIGHT: 'Left',
    SLOT_LENGTH_RIGHT: 'Left',
}


def is_annotation_element(element):
    """Return True when element is an annotation-like dependent element."""
    if not element:
        return False

    if isinstance(element, DB.IndependentTag):
        return True

    mra_type = getattr(DB, "MultiReferenceAnnotation", None)
    if mra_type and isinstance(element, mra_type):
        return True

    cat = element.Category
    return bool(cat and cat.CategoryType == DB.CategoryType.Annotation)


def get_annotation_dependents(host_element):
    """Collect annotation dependents attached to the host element."""
    if not host_element:
        return set()

    dependent_ids = set()
    try:
        raw_dep_ids = host_element.GetDependentElements(None)
    except Exception:
        raw_dep_ids = []

    for dep_id in raw_dep_ids or []:
        dep_elem = doc.GetElement(dep_id)
        if is_annotation_element(dep_elem):
            dependent_ids.add(dep_id)

    return dependent_ids


selected_ids = list(uidoc.Selection.GetElementIds())
if not selected_ids:
    script.exit()

host_elements = [doc.GetElement(eid) for eid in selected_ids]
annotation_ids = set()
for host in host_elements:
    annotation_ids.update(get_annotation_dependents(host))

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
        or get_bip_string(elem_type, "SYMBOL_FAMILY_NAME_PARAM")
        or get_bip_string(elem_type, "ALL_MODEL_FAMILY_NAME")
        or getattr(elem_type, "FamilyName", None)
    )
    type_name = (
        get_param_string(elem_type, "Type Name")
        or get_bip_string(elem_type, "SYMBOL_NAME_PARAM")
        or get_bip_string(elem_type, "ALL_MODEL_TYPE_NAME")
        or get_param_string(elem_type, "Type")
        or getattr(elem_type, "Name", None)
    )
    if family_name and type_name:
        return family_name, type_name
    return None, None


# Build lookup: (family_name, type_name) -> ElementTypeId
tuple_to_type_id = {}
all_element_types = DB.FilteredElementCollector(doc).OfClass(DB.ElementType)
for elem_type in all_element_types.ToElements():
    family_name, type_name = get_type_identity(elem_type)
    if family_name and type_name:
        tuple_to_type_id[(family_name, type_name)] = elem_type.Id


# Build lookup: source tag type id -> target tag type id
# Dictionary values are treated as target TYPE names within the same family.
type_swap_map = {}
for source_slot, target_type_name in dictonary.items():
    source_candidates = DEFAULT_TAG_SLOT_CANDIDATES.get(source_slot, [])
    if not source_candidates:
        continue

    for source_tag in source_candidates:
        if not isinstance(source_tag, tuple):
            continue
        if source_tag not in tuple_to_type_id:
            continue

        source_family_name, _ = source_tag
        source_type_id = tuple_to_type_id[source_tag]
        target_tuple = (source_family_name, target_type_name)
        target_type_id = tuple_to_type_id.get(target_tuple)

        if target_type_id:
            type_swap_map[source_type_id] = target_type_id


swapped_count = 0
with revit.Transaction("Switch tags based on dictionary"):
    for ann_id in annotation_ids:
        ann = doc.GetElement(ann_id)
        if not isinstance(ann, DB.IndependentTag):
            continue

        current_tag_id = ann.GetTypeId()
        if current_tag_id not in type_swap_map:
            continue

        target_tag_id = type_swap_map[current_tag_id]
        if target_tag_id == current_tag_id:
            continue

        ann.ChangeTypeId(target_tag_id)
        swapped_count += 1
