# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import revit, script
from Autodesk.Revit.DB import (
    BuiltInCategory,
    FilteredElementCollector,
    IndependentTag,
    Reference,
    TagMode,
    TagOrientation,
    Transaction,
    XYZ,
)
from tagging.revit_tagging import RevitTagging
from tagging.tag_config import (
    DEFAULT_TAG_SLOT_CANDIDATES,
    SLOT_CONDENSER,
    SLOT_CONDENSER_NOTE,
    SLOT_FAN,
    SLOT_FAN_NOTE,
    SLOT_HEAT_PUMP,
    SLOT_HEAT_PUMP_NOTE,
    SLOT_HEATER,
    SLOT_HEATER_NOTE,
    SLOT_HOOD,
    SLOT_HOOD_NOTE,
    SLOT_HUMIDIFIER,
    SLOT_HUMIDIFIER_NOTE,
    SLOT_SPLIT,
    SLOT_SPLIT_NOTE,
    SLOT_UNIT,
    SLOT_UNIT_NOTE,
    SLOT_VALVE,
    SLOT_VALVE_NOTE,
    SLOT_VRF,
    SLOT_VRF_NOTE,
)
from config.parameters_registry import BBM_SUBJECT, PYT_NOTE_0

# Button display information
# =================================================
__title__ = "Tag Equipment"
__doc__ = """

"""

# Subject -> equipment slot mapping.
SUBJECT_SLOT_MAP = {
    "condenser": SLOT_CONDENSER,
    "fan": SLOT_FAN,
    "heat pump": SLOT_HEAT_PUMP,
    "heater": SLOT_HEATER,
    "hood": SLOT_HOOD,
    "humidifier": SLOT_HUMIDIFIER,
    "split": SLOT_SPLIT,
    "unit": SLOT_UNIT,
    "valve": SLOT_VALVE,
    "vrf": SLOT_VRF,
}

SUBJECT_SLOT_MAP_NOTE = {
    "condenser": SLOT_CONDENSER_NOTE,
    "fan": SLOT_FAN_NOTE,
    "heat pump": SLOT_HEAT_PUMP_NOTE,
    "heater": SLOT_HEATER_NOTE,
    "hood": SLOT_HOOD_NOTE,
    "humidifier": SLOT_HUMIDIFIER_NOTE,
    "split": SLOT_SPLIT_NOTE,
    "unit": SLOT_UNIT_NOTE,
    "valve": SLOT_VALVE_NOTE,
    "vrf": SLOT_VRF_NOTE,
}

SUBJECT_PARAMETER_NAMES = (
    BBM_SUBJECT,
)
TARGET_CATEGORY = BuiltInCategory.OST_MechanicalEquipment

output = script.get_output()
doc = revit.doc
view = revit.active_view
tagger = RevitTagging(doc, view)


def _norm(text):
    return (text or '').strip().lower()


def _get_param(elem, *param_names):
    for param_name in param_names:
        if not param_name:
            continue
        try:
            param = elem.LookupParameter(param_name)
            if param:
                return param
        except Exception:
            continue
    return None


def _get_param_text(elem, *param_names):
    param = _get_param(elem, *param_names)
    if not param:
        return ''

    try:
        value = param.AsString()
    except Exception:
        value = None

    if not value:
        try:
            value = param.AsValueString()
        except Exception:
            value = None

    return (value or '').strip()


def _subject_slot(subject_text, has_note=False):
    subject_code = _norm(subject_text)
    if not subject_code:
        return None

    if has_note:
        return SUBJECT_SLOT_MAP_NOTE.get(subject_code)

    return SUBJECT_SLOT_MAP.get(subject_code)


def _resolve_tag_candidates(slot_name):
    resolved = []
    seen = set()
    configured_candidates = DEFAULT_TAG_SLOT_CANDIDATES.get(slot_name) or []

    for family_name, type_name in configured_candidates:
        key = (_norm(family_name), _norm(type_name))
        if key in seen:
            continue
        seen.add(key)

        try:
            tag_symbol = tagger.get_label_exact(family_name, type_name)
        except Exception:
            continue

        resolved.append((family_name, type_name, tag_symbol))

    return resolved


def _center_point(elem):
    try:
        location = elem.Location
        if hasattr(location, 'Point') and location.Point:
            return location.Point
    except Exception:
        pass

    bbox = elem.get_BoundingBox(view) if view else None
    if not bbox:
        return None

    min_pt = bbox.Min
    max_pt = bbox.Max
    return XYZ(
        (min_pt.X + max_pt.X) / 2.0,
        (min_pt.Y + max_pt.Y) / 2.0,
        (min_pt.Z + max_pt.Z) / 2.0,
    )


def _build_existing_tag_map():
    existing = {}
    tags = (
        FilteredElementCollector(doc, view.Id)
        .OfClass(IndependentTag)
        .ToElements()
    )

    for tag in tags:
        try:
            tag_type = doc.GetElement(tag.GetTypeId())
            family_name, type_name, _ = tagger._tag_pool(tag_type)
            tag_key = (_norm(family_name), _norm(type_name))
            if not all(tag_key):
                continue

            tagged_ids = []
            try:
                tagged_ids = list(tag.GetTaggedLocalElementIds() or [])
            except Exception:
                pass

            if not tagged_ids:
                try:
                    tagged_id = tag.TaggedLocalElementId
                    if tagged_id:
                        tagged_ids = [tagged_id]
                except Exception:
                    pass

            for tagged_id in tagged_ids:
                try:
                    elem_id = tagged_id.IntegerValue
                except Exception:
                    continue
                existing.setdefault(elem_id, set()).add(tag_key)
        except Exception:
            continue

    return existing


def _placed_tag_matches(tag, expected_family_name, expected_type_name):
    if tag is None:
        return False

    try:
        tag_type = doc.GetElement(tag.GetTypeId())
    except Exception:
        tag_type = None

    if tag_type is None:
        return False

    family_name, type_name, _ = tagger._tag_pool(tag_type)
    return (
        _norm(family_name) == _norm(expected_family_name)
        and _norm(type_name) == _norm(expected_type_name)
    )


def _get_tag_family_type(tag):
    if tag is None:
        return '', ''

    try:
        tag_type = doc.GetElement(tag.GetTypeId())
    except Exception:
        tag_type = None

    if tag_type is None:
        return '', ''

    family_name, type_name, _ = tagger._tag_pool(tag_type)
    return family_name, type_name


def _get_valid_tag_type_labels(tag, max_items=8):
    labels = []
    seen = set()

    if tag is None:
        return labels

    try:
        valid_type_ids = list(tag.GetValidTypes() or [])
    except Exception:
        valid_type_ids = []

    for valid_id in valid_type_ids:
        try:
            valid_type = doc.GetElement(valid_id)
        except Exception:
            valid_type = None

        if valid_type is None:
            continue

        family_name, type_name, _ = tagger._tag_pool(valid_type)
        label = '{} : {}'.format(family_name, type_name).strip()
        label_key = _norm(label)
        if not label_key or label_key in seen:
            continue

        seen.add(label_key)
        labels.append(label)
        if len(labels) >= max_items:
            break

    return labels


def _tag_symbol_category_id(tag_symbol):
    if tag_symbol is None:
        return None

    try:
        category = tag_symbol.Category
        if category and getattr(category, 'Id', None):
            return category.Id.IntegerValue
    except Exception:
        pass

    return None


def _placement_modes_for_tag(tag_symbol):
    symbol_category_id = _tag_symbol_category_id(tag_symbol)
    multi_category_id = int(BuiltInCategory.OST_MultiCategoryTags)

    if symbol_category_id == multi_category_id:
        return (
            TagMode.TM_ADDBY_MULTICATEGORY,
            TagMode.TM_ADDBY_CATEGORY,
        )

    return (
        TagMode.TM_ADDBY_CATEGORY,
        TagMode.TM_ADDBY_MULTICATEGORY,
    )


def _place_requested_tag(element_or_ref, tag_candidates, point_xyz):
    if element_or_ref is None:
        return None

    target = getattr(element_or_ref, 'element', element_or_ref)
    ref = target if isinstance(target, Reference) else Reference(target)
    valid_labels_by_mode = []

    for expected_family_name, expected_type_name, tag_symbol in tag_candidates:
        for tag_mode in _placement_modes_for_tag(tag_symbol):
            try:
                new_tag = IndependentTag.Create(
                    doc,
                    view.Id,
                    ref,
                    False,
                    tag_mode,
                    TagOrientation.Horizontal,
                    point_xyz,
                )
            except Exception:
                continue

            try:
                compatible_id = tagger._find_compatible_tag_type_id(new_tag, tag_symbol)
                if compatible_id is None:
                    valid_labels = _get_valid_tag_type_labels(new_tag)
                    if valid_labels:
                        valid_labels_by_mode.extend(valid_labels)
                    doc.Delete(new_tag.Id)
                    continue

                new_tag.ChangeTypeId(compatible_id)

                if _placed_tag_matches(new_tag, expected_family_name, expected_type_name):
                    return new_tag, expected_family_name, expected_type_name

                valid_labels = _get_valid_tag_type_labels(new_tag)
                if valid_labels:
                    valid_labels_by_mode.extend(valid_labels)
                doc.Delete(new_tag.Id)
            except Exception:
                try:
                    doc.Delete(new_tag.Id)
                except Exception:
                    pass

    deduped = []
    seen = set()
    for label in valid_labels_by_mode:
        label_key = _norm(label)
        if not label_key or label_key in seen:
            continue
        seen.add(label_key)
        deduped.append(label)

    if deduped:
        tagger.last_place_tag_failure = 'Requested tag type is not valid for this element; valid {}'.format(
            ', '.join(deduped))
    else:
        tagger.last_place_tag_failure = 'Requested tag type is not valid for this element'

    return None


def _candidate_keys(tag_candidates):
    return {
        (_norm(family_name), _norm(type_name))
        for family_name, type_name, _ in tag_candidates
    }


def _is_target_equipment(elem):
    if elem is None:
        return False

    try:
        category = elem.Category
        if not category or not getattr(category, 'Id', None):
            return False
        return category.Id.IntegerValue == int(TARGET_CATEGORY)
    except Exception:
        return False


def _collect_equipment_in_view():
    equipment = []
    seen_ids = set()

    def _append(elem):
        if not _is_target_equipment(elem):
            return

        try:
            elem_id = elem.Id.IntegerValue
        except Exception:
            return

        if elem_id in seen_ids:
            return

        seen_ids.add(elem_id)
        equipment.append(elem)

    for elem in (
        FilteredElementCollector(doc, view.Id)
        .WhereElementIsNotElementType()
        .ToElements()
    ):
        _append(elem)

    if equipment:
        return equipment

    for elem in (
        FilteredElementCollector(doc, view.Id)
        .OfCategory(TARGET_CATEGORY)
        .WhereElementIsNotElementType()
        .ToElements()
    ):
        _append(elem)

    return equipment


equipment_in_view = _collect_equipment_in_view()

if not equipment_in_view:
    output.print_md('## No mechanical equipment found in the active view.')
    script.exit()

existing_tag_map = _build_existing_tag_map()

placed = []
already_tagged = []
missing_subject = []
unmapped_subject = []
failed = []

t = Transaction(doc, 'Tag Equipment')
t.Start()
try:
    for elem in equipment_in_view:
        try:
            subject_text = _get_param_text(elem, *SUBJECT_PARAMETER_NAMES)
            if not subject_text:
                missing_subject.append(elem)
                continue

            has_note = bool(_get_param_text(elem, PYT_NOTE_0))
            slot_name = _subject_slot(subject_text, has_note)
            if not slot_name:
                unmapped_subject.append((elem, subject_text))
                continue

            tag_candidates = _resolve_tag_candidates(slot_name)
            if not tag_candidates:
                failed.append(
                    (
                        elem,
                        'Missing tag type for slot {}'.format(slot_name),
                    )
                )
                continue
            elem_id = elem.Id.IntegerValue

            if existing_tag_map.get(elem_id, set()) & _candidate_keys(tag_candidates):
                already_tagged.append(elem)
                continue

            tag_point = _center_point(elem)
            if not tag_point:
                failed.append((elem, 'Unable to determine tag point'))
                continue

            placed_result = _place_requested_tag(
                elem,
                tag_candidates,
                tag_point,
            )
            if not placed_result:
                failed.append(
                    (
                        elem,
                        tagger.last_place_tag_failure or 'Unable to place tag',
                    )
                )
                continue

            new_tag, family_name, type_name = placed_result
            tag_key = (_norm(family_name), _norm(type_name))

            placed.append(elem)
            existing_tag_map.setdefault(elem_id, set()).add(tag_key)
        except Exception as ex:
            failed.append((elem, str(ex)))

    t.Commit()
except Exception:
    t.RollBack()
    raise

output.print_md(
    '## Equipment Tagging: placed {}, already tagged {}, missing subject {}, unmapped {}, failed {}'.format(
        len(placed),
        len(already_tagged),
        len(missing_subject),
        len(unmapped_subject),
        len(failed),
    )
)

if unmapped_subject:
    output.print_md('### Unmapped Subject Codes')
    for elem, subject_text in unmapped_subject:
        output.print_md('- ID {}: {}'.format(output.linkify(elem.Id), subject_text))

if failed:
    output.print_md('### Failed Elements')
    for elem, reason in failed:
        output.print_md('- ID {}: {}'.format(output.linkify(elem.Id), reason))
