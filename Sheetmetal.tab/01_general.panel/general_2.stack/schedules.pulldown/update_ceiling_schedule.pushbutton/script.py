# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, FilteredElementCollector, XYZ, Transaction, StorageType
from pyrevit import revit, script

# Button info
# ======================================================================
__title__ = 'Schedule Ceiling Data'
__doc__ = '''
Lists all ceilings in the project with room/space and height data.
'''

# Variables
# ======================================================================

doc = revit.doc
view = revit.active_view
output = script.get_output()

csv_columns = ['room_name', 'label', 'ceiling_type', 'ceiling_height']
DEBUG_ROOM_NUMBERS = {'3228'}
CEILING_SLOT_INDEXES = [0, 1, 2, 3, 4]


def safe_as_string(param):
    if not param:
        return None
    try:
        text = param.AsString()
        if text:
            return text
    except Exception:
        pass
    try:
        text = param.AsValueString()
        if text:
            return text
    except Exception:
        pass
    return None


def read_type_name(ceiling, ceiling_type):
    if ceiling_type:
        try:
            type_name = getattr(ceiling_type, 'Name', None)
            if type_name:
                return type_name
        except Exception:
            pass

    for candidate in ['Type Name', 'Type']:
        try:
            text = safe_as_string(ceiling.LookupParameter(candidate))
            if text:
                return text
        except Exception:
            pass

    return 'N/A'


def normalize_ceiling_bucket(type_name):
    text = (type_name or '').upper()
    if not text:
        return 'UNKNOWN'
    if 'OTS' in text or 'OPEN TO STRUCTURE' in text:
        return 'OTS'
    if 'GWB' in text or 'GYPSUM' in text:
        return 'GWB'
    if 'ACT' in text or 'ACOUST' in text:
        return 'ACT'
    return 'OTHER'


def get_probe_point(element):
    try:
        location = getattr(element, 'Location', None)
        if location and hasattr(location, 'Point') and location.Point:
            return location.Point
    except Exception:
        pass

    try:
        bbox = element.get_BoundingBox(None)
        if bbox and bbox.Min and bbox.Max:
            return (bbox.Min + bbox.Max) * 0.5
    except Exception:
        pass

    return None


def resolve_room_or_space(point):
    if not point:
        return None, None, 'NONE'

    probes = [point]
    try:
        probes.append(point - XYZ(0, 0, 0.1))
    except Exception:
        pass

    for probe in probes:
        try:
            room = doc.GetRoomAtPoint(probe)
            if room:
                return room, 'ROOM', 'HOST'
        except Exception:
            pass

    for probe in probes:
        try:
            space = doc.GetSpaceAtPoint(probe)
            if space:
                return space, 'SPACE', 'HOST'
        except Exception:
            pass

    return None, None, 'NONE'


def read_name_number(spatial_elem):
    if not spatial_elem:
        return 'N/A', 'N/A'

    name = None
    number = None

    # Prefer built-in parameters because .Name/.Number can be inconsistent
    # on some Room/Space API objects.
    builtin_name_params = [
        BuiltInParameter.ROOM_NAME,
        BuiltInParameter.SPACE_ASSOC_ROOM_NAME,
    ]
    builtin_number_params = [
        BuiltInParameter.ROOM_NUMBER,
        BuiltInParameter.SPACE_ASSOC_ROOM_NUMBER,
    ]

    for bip in builtin_name_params:
        try:
            text = safe_as_string(spatial_elem.get_Parameter(bip))
            if text:
                name = text
                break
        except Exception:
            pass

    for bip in builtin_number_params:
        try:
            text = safe_as_string(spatial_elem.get_Parameter(bip))
            if text:
                number = text
                break
        except Exception:
            pass

    # Shared parameter / localized-name fallback.
    if not name:
        for pname in ['Name', 'Room Name', 'Space Name']:
            try:
                text = safe_as_string(spatial_elem.LookupParameter(pname))
                if text:
                    name = text
                    break
            except Exception:
                pass

    if not number:
        for pname in ['Number', 'Room Number', 'Space Number']:
            try:
                text = safe_as_string(spatial_elem.LookupParameter(pname))
                if text:
                    number = text
                    break
            except Exception:
                pass

    # Last resort dynamic properties.
    try:
        if not name:
            name = getattr(spatial_elem, 'Name', None)
    except Exception:
        pass

    try:
        if not number:
            number = getattr(spatial_elem, 'Number', None)
    except Exception:
        pass

    return (name or 'N/A'), (number or 'N/A')


def read_height(ceiling):
    # Try multiple parameter name variations
    param_names = [
        'Height Offset From Level [Instance]',
        'Height Offset From Level',
        'Offset From Level',
        'Height',
    ]
    for param_name in param_names:
        try:
            param = ceiling.LookupParameter(param_name)
            text = safe_as_string(param)
            if text:
                numeric_value = None
                try:
                    numeric_value = param.AsDouble()
                except Exception:
                    numeric_value = None

                # Filter out negative heights (invalid)
                if numeric_value is not None and numeric_value < 0:
                    continue
                return text, numeric_value
        except Exception:
            pass
    return None, None


def room_sort_key(room_number):
    text = room_number or ''
    stripped = text.lstrip('0') or '0'
    if stripped.isdigit():
        return 0, int(stripped), text
    return 1, text


def is_non_negative_height(height_value):
    return height_value is not None and height_value >= 0


def should_replace_selected(existing, candidate):
    if not existing:
        return True

    existing_height = existing.get('height_value')
    candidate_height = candidate.get('height_value')
    existing_valid = is_non_negative_height(existing_height)
    candidate_valid = is_non_negative_height(candidate_height)

    if candidate_valid and not existing_valid:
        return True

    if candidate_valid and existing_valid:
        return candidate_height < existing_height

    return False


def make_room_key(room_number, room_name=None):
    return room_number or 'N/A'


def should_debug_room(room_number):
    return (room_number or '') in DEBUG_ROOM_NUMBERS


def set_param_value(param, text_value=None, double_value=None):
    if not param:
        return False, 'missing parameter'

    try:
        if param.IsReadOnly:
            return False, 'read-only'
    except Exception:
        pass

    try:
        storage_type = param.StorageType
    except Exception as exc:
        return False, 'no storage type: {}'.format(exc)

    try:
        if storage_type == StorageType.String:
            if text_value is None:
                return False, 'string param needs text value'
            param.Set(text_value)
            return True, 'set string'

        if storage_type == StorageType.Double:
            if double_value is None:
                return False, 'double param needs numeric value'
            param.Set(double_value)
            return True, 'set double'

        if storage_type == StorageType.Integer:
            return False, 'integer storage not supported'

        if storage_type == StorageType.ElementId:
            return False, 'elementId storage not supported'

        return False, 'unknown storage type'
    except Exception as exc:
        return False, str(exc)


def clear_param_value(param):
    if not param:
        return False, 'missing parameter'

    try:
        if param.IsReadOnly:
            return False, 'read-only'
    except Exception:
        pass

    try:
        storage_type = param.StorageType
    except Exception as exc:
        return False, 'no storage type: {}'.format(exc)

    try:
        if storage_type == StorageType.String:
            param.Set('')
            return True, 'cleared string'

        if storage_type == StorageType.Double:
            param.Set(0.0)
            return True, 'cleared double'

        return False, 'clear not supported'
    except Exception as exc:
        return False, str(exc)


def ranked_room_ceilings(candidates):
    # Sort by non-negative numeric height first (low -> high), then keep unique type/height pairs.
    def sort_key(item):
        hv = item.get('height_value')
        if hv is not None and hv >= 0:
            return (0, hv, item.get('height_text') or '', item.get('bucket') or '')
        return (1, float('inf'), item.get('height_text') or '', item.get('bucket') or '')

    unique = []
    seen = set()
    for item in sorted(candidates, key=sort_key):
        sig = (item.get('bucket') or '', item.get('height_text') or '')
        if sig in seen:
            continue
        seen.add(sig)
        unique.append(item)
    return unique


ceilings = list((FilteredElementCollector(doc)
                 .OfCategory(BuiltInCategory.OST_Ceilings)
                 .WhereElementIsNotElementType()
                 .ToElements()))

if not ceilings:
    output.print_md('## No ceilings found in project.')
    script.exit()

# Collect all rooms in the document
rooms = list((FilteredElementCollector(doc)
             .OfCategory(BuiltInCategory.OST_Rooms)
             .WhereElementIsNotElementType()
             .ToElements()))

output.print_md('# Ceiling Data Pull')
output.print_md('---')
output.print_md('Writing ceiling data to custom parameters...')
output.print_md('')

update_count = 0
ceiling_data = []  # Collect all ceiling data first
rooms_with_ceilings = set()
selected_ceiling_by_room_key = {}
ceilings_by_room_key = {}

with Transaction(doc, 'Update Ceiling Parameters') as txn:
    txn.Start()
    for idx, ceiling in enumerate(ceilings, start=1):
        ceiling_type = None
        try:
            ceiling_type = doc.GetElement(ceiling.GetTypeId())
        except Exception:
            ceiling_type = None

        type_name = read_type_name(ceiling, ceiling_type)
        bucket = normalize_ceiling_bucket(type_name)
        height_text, height_value = read_height(ceiling)

        probe = get_probe_point(ceiling)
        spatial, spatial_kind, _ = resolve_room_or_space(probe)
        room_name, room_number = read_name_number(spatial)
        room_key = make_room_key(room_number)

        # Store ceiling data for sorting
        ceiling_data.append({
            'idx': idx,
            'ceiling_id': ceiling.Id,
            'ceiling_id_value': ceiling.Id.IntegerValue,
            'spatial': spatial,
            'spatial_kind': spatial_kind,
            'room_key': room_key,
            'room_number': room_number,
            'room_name': room_name,
            'type_name': type_name,
            'bucket': bucket,
            'height_text': height_text,
            'height_value': height_value,
        })

        if spatial:
            rooms_with_ceilings.add(room_key)
            if room_key not in ceilings_by_room_key:
                ceilings_by_room_key[room_key] = []
            ceilings_by_room_key[room_key].append(ceiling_data[-1])
            existing = selected_ceiling_by_room_key.get(room_key)
            if should_replace_selected(existing, ceiling_data[-1]):
                selected_ceiling_by_room_key[room_key] = ceiling_data[-1]

    selected_ceiling_ids = set(item['ceiling_id_value']
                               for item in selected_ceiling_by_room_key.values())

    for room in rooms:
        room_name, room_number = read_name_number(room)
        room_key = make_room_key(room_number)
        selected = selected_ceiling_by_room_key.get(room_key)
        try:
            p_room_name = room.LookupParameter('_UMI_PYT_RoomName')
            p_room_number = room.LookupParameter('_UMI_PYT_RoomNumber')
            p_ceiling_type = room.LookupParameter('_UMI_PYT_CeilingType')
            p_ceiling_height = room.LookupParameter('_UMI_PYT_CeilingHeight')
            slot_type_params = [room.LookupParameter(
                '_UMI_PYT_CeilingType{}'.format(i)) for i in CEILING_SLOT_INDEXES]
            slot_height_params = [room.LookupParameter('_UMI_PYT_CeilingHeight{}'.format(i))
                                  for i in CEILING_SLOT_INDEXES]

            if p_room_name:
                p_room_name.Set(room_name)
            if p_room_number:
                p_room_number.Set(room_number)

            before_height = safe_as_string(p_ceiling_height)
            height_write_status = 'not attempted'

            if selected and is_non_negative_height(selected['height_value']):
                if p_ceiling_type:
                    p_ceiling_type.Set(selected['bucket'])
                if p_ceiling_height:
                    ok, msg = set_param_value(
                        p_ceiling_height,
                        text_value=selected['height_text'],
                        double_value=selected['height_value'],
                    )
                    height_write_status = ('OK: ' if ok else 'FAIL: ') + msg
            else:
                if p_ceiling_type:
                    p_ceiling_type.Set('OTS')
                if selected and not is_non_negative_height(selected['height_value']):
                    height_write_status = 'skipped: selected ceiling has no valid height'
                else:
                    height_write_status = 'skipped: no selected ceiling'

            ranked = ranked_room_ceilings(
                ceilings_by_room_key.get(room_key, []))
            for slot_idx in range(len(CEILING_SLOT_INDEXES)):
                p_type_slot = slot_type_params[slot_idx]
                p_height_slot = slot_height_params[slot_idx]
                slot_item = ranked[slot_idx] if slot_idx < len(
                    ranked) else None

                if slot_item and is_non_negative_height(slot_item['height_value']):
                    if p_type_slot:
                        p_type_slot.Set(slot_item['bucket'])
                    if p_height_slot:
                        set_param_value(
                            p_height_slot,
                            text_value=slot_item['height_text'],
                            double_value=slot_item['height_value'],
                        )
                else:
                    if p_type_slot:
                        p_type_slot.Set('OTS' if slot_item else '')
                    if p_height_slot:
                        clear_param_value(p_height_slot)

            if should_debug_room(room_number):
                after_height = safe_as_string(p_ceiling_height)
                storage_name = 'N/A'
                try:
                    storage_name = str(
                        p_ceiling_height.StorageType) if p_ceiling_height else 'MISSING'
                except Exception:
                    storage_name = 'UNKNOWN'
                output.print_md(
                    '### DEBUG WRITE: Room {} | Room Id {} | Selected Ceiling Id {} | Selected Height {} | Before {} | After {} | Storage {} | Write {}'.format(
                        room_number,
                        output.linkify(
                            room.Id),
                        output.linkify(
                            selected['ceiling_id']) if selected else 'NONE',
                        selected['height_text'] if selected and selected['height_text'] is not None else 'N/A',
                        before_height if before_height is not None else 'N/A',
                        after_height if after_height is not None else 'N/A',
                        storage_name,
                        height_write_status,
                    ))
                output.print_md('Debug Room Name: {}'.format(room_name))

            update_count += 1
        except Exception as e:
            output.print_md(
                'Error updating room {}: {}'.format(room_number, str(e)))

    # Sort ceiling data by room number for easier reading
    try:
        ceiling_data.sort(key=lambda x: (
            room_sort_key(x['room_number']), x['idx']))
    except Exception:
        ceiling_data.sort(key=lambda x: x['room_number'])

    # Print sorted ceiling data
    for item in ceiling_data:
        selected_marker = ' [SELECTED]' if item['ceiling_id_value'] in selected_ceiling_ids else ''
        output.print_md(
            '### {:03d}: ID {} | {} {} | Type: {} | Bucket: {} | Height: {}{}'.format(
                item['idx'],
                output.linkify(item['ceiling_id']),
                item['spatial_kind'] or 'NO-ROOM',
                item['room_number'],
                item['type_name'],
                item['bucket'],
                item['height_text'] if item['height_text'] is not None else 'N/A',
                selected_marker,
            )
        )
        output.print_md('Room/Space Name: {}'.format(item['room_name']))

    # Process rooms without ceilings
    output.print_md('')
    output.print_md('---')
    output.print_md('Processing rooms without ceilings...')
    output.print_md('')

    rooms_without_ceilings = []
    for room in rooms:
        room_name, room_number = read_name_number(room)
        room_key = make_room_key(room_number)
        if room_key not in rooms_with_ceilings:
            rooms_without_ceilings.append({
                'room': room,
                'room_name': room_name,
                'room_number': room_number,
            })

    rooms_without_ceilings.sort(
        key=lambda item: room_sort_key(item['room_number']))

    for item in rooms_without_ceilings:
        room_name = item['room_name']
        room_number = item['room_number']
        output.print_md(
            '### NO CEILING: Room {} | Type: OTS'.format(room_number))
        output.print_md('Room Name: {}'.format(room_name))

    txn.Commit()

output.print_md('---')
output.print_md('**Total ceilings: {}**'.format(len(ceilings)))
output.print_md('**Total rooms: {}**'.format(len(rooms)))
output.print_md('**Parameters updated: {}**'.format(update_count))
