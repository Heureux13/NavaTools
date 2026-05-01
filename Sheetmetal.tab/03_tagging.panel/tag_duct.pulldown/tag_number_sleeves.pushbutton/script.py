# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import revit, script
from config.parameters_registry import *
from tagging.tag_config import DEFAULT_TAG_SLOT_CANDIDATES, SLOT_NUMBER_SLEEVE
from Autodesk.Revit.DB import (
    BuiltInCategory,
    ElementId,
    ElementTransformUtils,
    FabricationPart,
    FilteredElementCollector,
    IndependentTag,
    Line,
    StorageType,
    XYZ,
)
import math
import re
try:
    from tagging.revit_tagging import RevitTagging  # type: ignore[reportMissingImports]
except ImportError:
    from tagging.revit_tagging import RevitTagging

# Button info
# ======================================================================
__title__ = 'Tag Pen Sleeves'
__doc__ = '''
Will calculte the size for sleeves, number them, and tag them. so long as they have the _type paramter value of "sleeve"
'''

# Variables
# ======================================================================
uidoc = getattr(revit, 'uidoc', None)
doc = getattr(revit, 'doc', None)
view = getattr(revit, 'active_view', None)

if uidoc is None or doc is None or view is None:
    revit_host = globals().get('__revit__')
    if revit_host is not None:
        uidoc = uidoc or revit_host.ActiveUIDocument
        if uidoc is not None:
            doc = doc or uidoc.Document
            view = view or uidoc.ActiveView

if uidoc is None:
    raise RuntimeError('Unable to access the active Revit UI document.')
if doc is None:
    raise RuntimeError('Unable to access the active Revit document.')
if view is None:
    raise RuntimeError('Unable to access the active Revit view.')

output = script.get_output()

TYPE_PARAM = PYT_SLEEVE_VALUE
NUMBER_PARAM = PYT_NUMBER_SLEEVE
SLEEVE_VALUE = 'sleeve'
TAG_PARAM = PYT_NUMBER_SLEEVE
ANNOTATION_TAG_CANDIDATES = list(DEFAULT_TAG_SLOT_CANDIDATES.get(SLOT_NUMBER_SLEEVE, []))
SIZE = RVT_SIZE
CONNECTOR = NDBS_CONNECTOR0_END_CONDITION
WRAP = RVT_INSULATION_SPECIFICATION
FAMILY = NDBS_FAMILY
DEFAULT_CLEARANCE = 1
SLEEVE_OPENING = PYT_SLEEVE_OPENING

CONNECTOR_RULES = {
    'tdf': {
        'connector_add': 3.0,
        'use_wrap': False,
    },
    'tdc': {
        'connector_add': 3.0,
        'use_wrap': False,
    },
    's&d': {
        'connector_add': 0.0,
        'use_wrap': True,
    },
    's and d': {
        'connector_add': 0.0,
        'use_wrap': True,
    },
}

DEFAULT_CONNECTOR_RULE = {
    'connector_add': 0.0,
    'use_wrap': True,
}

MAX_WRAP_INCHES = 24.0


def _normalize_text(value):
    if value is None:
        return ''
    text = str(value).strip().lower()
    return re.sub(r'\s+', ' ', text)


NORMALIZED_CONNECTOR_RULES = {
    _normalize_text(key): value
    for key, value in CONNECTOR_RULES.items()
}


def _element_id_value(element_or_id):
    try:
        return element_or_id.Id.Value
    except AttributeError:
        pass

    try:
        return element_or_id.Id.IntegerValue
    except AttributeError:
        pass

    try:
        return element_or_id.Value
    except AttributeError:
        return element_or_id.IntegerValue


def _get_param_string(element, param_name):
    param = element.LookupParameter(param_name)
    if not param:
        return None

    try:
        value = param.AsString()
        if value is None:
            value = param.AsValueString()
        return value
    except Exception:
        return None


def _normalized_param_value(element, param_name):
    value = _get_param_string(element, param_name)
    return _normalize_text(value)


def _set_param_value(element, param_name, value):
    param = element.LookupParameter(param_name)
    if not param or param.IsReadOnly:
        return False

    try:
        if param.StorageType == StorageType.String:
            param.Set(str(value))
            return True
        if param.StorageType == StorageType.Integer:
            param.Set(int(value))
            return True
        if param.StorageType == StorageType.Double:
            param.Set(float(value))
            return True
    except Exception:
        return False

    return False


def _lookup_writable_param(element, param_name):
    param = element.LookupParameter(param_name)
    if not param:
        return None, 'missing'
    if param.IsReadOnly:
        return None, 'readonly'
    return param, None


def _parse_positive_int(value):
    if value is None:
        return None

    try:
        parsed = int(value)
        return parsed if parsed > 0 else None
    except Exception:
        pass

    text = str(value).strip()
    if not text:
        return None

    try:
        parsed = int(float(text))
        return parsed if parsed > 0 else None
    except Exception:
        pass

    match = re.search(r'\d+', text)
    if not match:
        return None

    try:
        parsed = int(match.group(0))
        return parsed if parsed > 0 else None
    except Exception:
        return None


def _get_positive_int_param(element, param_name):
    param = element.LookupParameter(param_name)
    if not param:
        return None

    try:
        if param.StorageType == StorageType.Integer:
            return _parse_positive_int(param.AsInteger())
        if param.StorageType == StorageType.Double:
            return _parse_positive_int(param.AsDouble())
    except Exception:
        pass

    value = _get_param_string(element, param_name)
    return _parse_positive_int(value)


def _eid_int(eid):
    if eid is None:
        return None

    host_id = getattr(eid, 'HostElementId', None)
    if host_id is not None:
        eid = host_id

    for attr in ('Value', 'IntegerValue'):
        try:
            value = getattr(eid, attr)
            if value is not None:
                return int(value)
        except Exception:
            pass

    try:
        return int(eid)
    except Exception:
        return None


def _collect_tagged_local_ids(tag):
    ids = []

    try:
        for tid in tag.GetTaggedLocalElementIds() or []:
            if tid and tid != ElementId.InvalidElementId:
                ids.append(tid)
    except Exception:
        pass

    try:
        tid = tag.TaggedLocalElementId
        if tid and tid != ElementId.InvalidElementId:
            ids.append(tid)
    except Exception:
        pass

    def _append_from_link_eid(link_eid):
        if not link_eid:
            return
        for attr in ('HostElementId', 'LinkedElementId', 'ElementId'):
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

    unique = []
    seen = set()
    for tid in ids:
        tid_int = _eid_int(tid)
        if tid_int is None or tid_int in seen:
            continue
        seen.add(tid_int)
        unique.append(tid)
    return unique


def _tag_family_name(tag):
    if tag is None:
        return ''

    try:
        tag_type = doc.GetElement(tag.GetTypeId())
    except Exception:
        tag_type = None

    if tag_type is None:
        return ''

    family = getattr(tag_type, 'Family', None)
    family_name = getattr(family, 'Name', '') if family is not None else ''
    return _normalize_text(family_name)


def _format_dim(value):
    rounded = round(float(value), 3)
    if abs(rounded - round(rounded)) < 1e-6:
        return str(int(round(rounded)))
    return ('{:.3f}'.format(rounded)).rstrip('0').rstrip('.')


def _parse_size_pair(raw_size):
    if not raw_size:
        return None
    text = str(raw_size).strip().lower()
    if not text:
        return None

    text = text.replace('×', 'x')
    match = re.search(r'([0-9]+(?:\.[0-9]+)?)\s*x\s*([0-9]+(?:\.[0-9]+)?)', text)
    if not match:
        nums = re.findall(r'[0-9]+(?:\.[0-9]+)?', text)
        if len(nums) == 1:
            # Round sizes like 14"ø are treated as diameter x diameter.
            diameter = float(nums[0])
            return diameter, diameter, True
        if len(nums) < 2:
            return None
        return float(nums[0]), float(nums[1]), ('ø' in text)

    return float(match.group(1)), float(match.group(2)), False


def _get_param_number(element, param_name, length_double_to_inches=False):
    param = element.LookupParameter(param_name)
    if not param:
        return None

    try:
        if param.StorageType == StorageType.Double:
            val = float(param.AsDouble())
            return val * 12.0 if length_double_to_inches else val
        if param.StorageType == StorageType.Integer:
            return float(param.AsInteger())
    except Exception:
        pass

    value = _get_param_string(element, param_name)
    if value is None:
        return None

    match = re.search(r'[0-9]+(?:\.[0-9]+)?', str(value))
    if not match:
        return None

    try:
        return float(match.group(0))
    except Exception:
        return None


def _build_sleeve_opening_value(element):
    size_raw = _get_param_string(element, SIZE)
    size_pair = _parse_size_pair(size_raw)
    if size_pair is None:
        return None

    width, height, is_round = size_pair

    family_value = _normalized_param_value(element, FAMILY)
    uses_connector_rules = (
        ('straight' in family_value) or
        ('spiral' in family_value)
    )
    if not uses_connector_rules:
        total_add = float(DEFAULT_CLEARANCE)
        opening_width = width + total_add
        opening_height = height + total_add
        if is_round:
            return '{}{}'.format(_format_dim(opening_width), chr(216))
        return '{}x{}'.format(_format_dim(opening_width), _format_dim(opening_height))

    connector_key = _normalized_param_value(element, CONNECTOR)
    rule = NORMALIZED_CONNECTOR_RULES.get(connector_key, DEFAULT_CONNECTOR_RULE)

    wrap_value = _get_param_number(element, WRAP, length_double_to_inches=True)
    # Some fabrication models return sentinel values for missing insulation.
    if wrap_value is None or wrap_value < 0.0 or wrap_value > MAX_WRAP_INCHES:
        wrap_value = 0.0

    total_add = float(DEFAULT_CLEARANCE) + float(rule['connector_add'])
    if bool(rule['use_wrap']):
        total_add += float(wrap_value) * 2.0

    opening_width = width + total_add
    opening_height = height + total_add
    if is_round:
        return '{}{}'.format(_format_dim(opening_width), chr(216))
    return '{}x{}'.format(_format_dim(opening_width), _format_dim(opening_height))


def _find_first_tag_symbol(tagger, candidates):
    for candidate in candidates:
        try:
            if isinstance(candidate, tuple):
                tag_symbol = tagger.get_label_exact(candidate[0], candidate[1])
            else:
                tag_symbol = tagger.get_label(candidate)
            if tag_symbol is not None:
                return tag_symbol, candidate
        except LookupError:
            continue
    return None, None


def _tag_point_for_element(element):
    try:
        location = getattr(element, 'Location', None)
        if location is not None and hasattr(location, 'Curve') and location.Curve:
            return location.Curve.Evaluate(0.5, True)
    except Exception:
        pass

    try:
        bbox = element.get_BoundingBox(view)
        if bbox is not None:
            return (bbox.Min + bbox.Max) / 2.0
    except Exception:
        pass

    return None


def _rotate_tag_to_element(tag, element):
    try:
        location = getattr(element, 'Location', None)
        if location is None or not hasattr(location, 'Curve') or not location.Curve:
            return False

        curve = location.Curve
        direction = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize()
        angle_radians = math.atan2(direction.Y, direction.X)

        head = tag.TagHeadPosition
        axis = Line.CreateBound(
            head,
            XYZ(head.X, head.Y, head.Z + 1.0)
        )
        ElementTransformUtils.RotateElement(doc, tag.Id, axis, angle_radians)
        return True
    except Exception:
        return False


all_ducts = list(
    FilteredElementCollector(doc)
    .OfClass(FabricationPart)
    .WhereElementIsNotElementType()
    .ToElements()
)

sleeves = []
missing_type = []

for element in all_ducts:
    category = getattr(element, 'Category', None)
    if category is None:
        continue
    if _element_id_value(category.Id) != int(BuiltInCategory.OST_FabricationDuctwork):
        continue

    type_value = _normalized_param_value(element, TYPE_PARAM)

    if not type_value:
        missing_type.append(element)
        continue

    if type_value == SLEEVE_VALUE:
        sleeves.append(element)

if not sleeves:
    output.print_md('## No sleeve elements found in the active view.')
    output.print_md('Fabrication ductwork checked: {}'.format(len(all_ducts)))
    if missing_type:
        output.print_md(
            'Elements missing `{}`: {}'.format(TYPE_PARAM, len(missing_type)))
    script.exit()

sleeves.sort(key=_element_id_value)

tagger = RevitTagging(doc, view)
annotation_tag_symbol, annotation_tag_name = _find_first_tag_symbol(
    tagger, ANNOTATION_TAG_CANDIDATES)
annotation_tag_family = ''
if annotation_tag_symbol is not None and annotation_tag_symbol.Family is not None:
    annotation_tag_family = (annotation_tag_symbol.Family.Name or '').strip().lower()

existing_annotation_tag_map = {}
if annotation_tag_symbol is not None:
    existing_annotation_tags = list(
        FilteredElementCollector(doc, view.Id)
        .OfClass(IndependentTag)
        .ToElements()
    )
    for existing_tag in existing_annotation_tags:
        try:
            if _tag_family_name(existing_tag) != annotation_tag_family:
                continue
            for tagged_id in _collect_tagged_local_ids(existing_tag):
                tagged_id_int = _eid_int(tagged_id)
                if tagged_id_int is None:
                    continue
                existing_annotation_tag_map.setdefault(tagged_id_int, []).append(existing_tag)
        except Exception:
            continue

numbered = []
already_numbered = []
failed = []
tagged = []
tag_failed = []
tag_param_missing = []
tag_param_readonly = []
annotations_placed = []
annotations_failed = []
annotations_already_tagged = []
opening_updated = []
opening_failed = []
opening_param_missing = []
opening_param_readonly = []
opening_no_value = []
size_parse_failed = []
has_tag_param = any(s.LookupParameter(TAG_PARAM) for s in sleeves)

existing_numbers = []
unnumbered_sleeves = []
for sleeve in sleeves:
    existing_number = _get_positive_int_param(sleeve, NUMBER_PARAM)
    if existing_number is None:
        unnumbered_sleeves.append(sleeve)
    else:
        existing_numbers.append(existing_number)
        already_numbered.append(sleeve)

next_number = (max(existing_numbers) if existing_numbers else 0) + 1
assigned_numbers = {}

with revit.Transaction('Number sleeve elements'):
    for element in unnumbered_sleeves:
        assigned_value = next_number
        next_number += 1

        if _set_param_value(element, NUMBER_PARAM, assigned_value):
            numbered.append(element)
            assigned_numbers[_element_id_value(element.Id)] = assigned_value
        else:
            failed.append(element)

    for element in sleeves:
        element_key = _element_id_value(element.Id)
        number_value = assigned_numbers.get(
            element_key,
            _get_positive_int_param(element, NUMBER_PARAM)
        )
        if number_value is None:
            tag_failed.append(element)
            continue

        if has_tag_param:
            current_tag_value = _get_positive_int_param(element, TAG_PARAM)
            if current_tag_value == number_value:
                tagged.append(element)
            else:
                _, tag_param_issue = _lookup_writable_param(element, TAG_PARAM)
                if tag_param_issue == 'missing':
                    tag_param_missing.append(element)
                elif tag_param_issue == 'readonly':
                    tag_param_readonly.append(element)
                elif _set_param_value(element, TAG_PARAM, number_value):
                    tagged.append(element)
                else:
                    tag_failed.append(element)

        opening_value = _build_sleeve_opening_value(element)
        if opening_value is None:
            raw_size = _get_param_string(element, SIZE)
            if _parse_size_pair(raw_size) is None:
                size_parse_failed.append(element)
            else:
                opening_no_value.append(element)
        else:
            _, opening_param_issue = _lookup_writable_param(element, SLEEVE_OPENING)
            if opening_param_issue == 'missing':
                opening_param_missing.append(element)
            elif opening_param_issue == 'readonly':
                opening_param_readonly.append(element)
            elif _set_param_value(element, SLEEVE_OPENING, opening_value):
                opening_updated.append(element)
            else:
                opening_failed.append(element)

        if annotation_tag_symbol is None or not annotation_tag_family:
            continue

        existing_for_elem = existing_annotation_tag_map.get(_eid_int(element.Id), [])
        if existing_for_elem:
            annotations_already_tagged.append(element)
            continue

        tag_point = _tag_point_for_element(element)
        if tag_point is None:
            annotations_failed.append(element)
            continue

        try:
            annotation_tag = tagger.place_tag(element, annotation_tag_symbol, tag_point)
            _rotate_tag_to_element(annotation_tag, element)
            annotations_placed.append(element)
            elem_id_int = _eid_int(element.Id)
            if elem_id_int is not None:
                existing_annotation_tag_map.setdefault(elem_id_int, []).append(annotation_tag)
        except Exception:
            annotations_failed.append(element)

output.print_md('## Sleeve numbering complete.')
output.print_md('Fabrication ductwork checked: {}'.format(len(all_ducts)))
output.print_md('Sleeves found: {}'.format(len(sleeves)))
output.print_md('Already numbered in `{}` (kept): {}'.format(NUMBER_PARAM, len(already_numbered)))
output.print_md('Newly numbered in `{}`: {}'.format(NUMBER_PARAM, len(numbered)))
if has_tag_param:
    output.print_md('Annotated in `{}`: {}'.format(TAG_PARAM, len(tagged)))
else:
    output.print_md('Parameter `{}` not found on sleeves; skipped parameter write.'.format(TAG_PARAM))
output.print_md('Updated `{}`: {}'.format(SLEEVE_OPENING, len(opening_updated)))
if annotation_tag_symbol is None:
    output.print_md(
        'No loaded annotation tag found for: {}'.format(', '.join(ANNOTATION_TAG_CANDIDATES))
    )
else:
    output.print_md(
        'Placed annotation tags (`{}`): {}'.format(annotation_tag_name, len(annotations_placed))
    )
    if annotations_already_tagged:
        output.print_md(
            'Already tagged with same family: {}'.format(len(annotations_already_tagged))
        )
    if annotations_failed:
        output.print_md('Failed to place annotation tag: {}'.format(len(annotations_failed)))

if missing_type:
    output.print_md('Elements missing `{}`: {}'.format(TYPE_PARAM, len(missing_type)))
if failed:
    output.print_md('Failed to write `{}`: {}'.format(NUMBER_PARAM, len(failed)))
if has_tag_param and tag_failed:
    output.print_md('Failed to write `{}`: {}'.format(TAG_PARAM, len(tag_failed)))
if has_tag_param and tag_param_missing:
    output.print_md('Missing parameter `{}`: {}'.format(TAG_PARAM, len(tag_param_missing)))
if has_tag_param and tag_param_readonly:
    output.print_md('Read-only parameter `{}`: {}'.format(TAG_PARAM, len(tag_param_readonly)))
if opening_failed:
    output.print_md('Failed to write `{}`: {}'.format(SLEEVE_OPENING, len(opening_failed)))
if opening_no_value:
    output.print_md('No computed opening value: {}'.format(len(opening_no_value)))
if size_parse_failed:
    output.print_md('Failed to parse `{}` size: {}'.format(SIZE, len(size_parse_failed)))
    size_failed_ids = [e.Id for e in size_parse_failed]
    output.print_md('Size parse failed elements: {}'.format(output.linkify(size_failed_ids)))
if opening_param_missing:
    output.print_md('Missing parameter `{}`: {}'.format(SLEEVE_OPENING, len(opening_param_missing)))
if opening_param_readonly:
    output.print_md('Read-only parameter `{}`: {}'.format(SLEEVE_OPENING, len(opening_param_readonly)))
