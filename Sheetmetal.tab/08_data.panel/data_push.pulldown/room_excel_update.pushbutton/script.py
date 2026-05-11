# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    ElementId,
    TextNote,
)
from pyrevit import revit, script
from System.Windows.Forms import DialogResult, OpenFileDialog
import clr
import csv
import codecs


# Button info
# ===================================================
__title__ = "Rooms - Update w/ Excel"
__doc__ = """
Lists all room tags in the active view.
"""

# Variables
# ==================================================
doc = revit.doc
view = revit.active_view
output = script.get_output()

# CSV mapping configuration
excel_layer_column_value = 'info - rooms'
column_room_number_name = ['label', 'space']
column_heigth_name = 'size'
column_ceiling_type_name = 'ceiling'
column_room_number_fallback_names = ['og label', 'space']
valid_ceiling_types = ['ACT', 'GWB', 'OTS']
default_ceiling_type = 'XXX'
default_ceiling_size = '09\'-00"'
default_second_tag_text = 'XXX 09\'-00"'

clr.AddReference('System.Windows.Forms')


def is_missing(value):
    return not value or str(value).strip() == "" or str(value) == "N/A"


def normalize_key(value):
    if value is None:
        return ''
    text = str(value).replace('\ufeff', '').strip().lower()
    return text


def normalize_header(value):
    text = normalize_key(value)
    text = text.replace('-', '_').replace(' ', '_')
    while '__' in text:
        text = text.replace('__', '_')
    return text.strip('_')


def normalize_ceiling_type(value):
    text = (value or '').strip().upper()
    if not text:
        return ''
    token = text.split(' ')[0]
    if token in valid_ceiling_types:
        return token
    return ''


def build_ceiling_text(ceiling_type, size_text):
    ct = normalize_ceiling_type(ceiling_type) or default_ceiling_type
    if ct == 'OTS':
        return 'OTS'
    sz = (size_text or '').strip() or default_ceiling_size
    return '{} {}'.format(ct, sz).strip()


def looks_like_valid_ceiling_text(text):
    if is_missing(text):
        return False
    first = str(text).strip().upper().split(' ')[0]
    return first in valid_ceiling_types


def pick_csv_file():
    dialog = OpenFileDialog()
    dialog.Title = 'Select CSV File with Room Data'
    dialog.Filter = 'CSV Files (*.csv)|*.csv|All Files (*.*)|*.*'
    dialog.Multiselect = False

    if dialog.ShowDialog() != DialogResult.OK:
        return None
    return dialog.FileName


def load_csv_rows(file_path):
    rows = []
    with codecs.open(file_path, 'r', 'utf-8-sig') as handle:
        reader = csv.reader(handle)
        headers = None

        for raw_row in reader:
            if not raw_row:
                continue

            if headers is None:
                headers = [normalize_header(h) for h in raw_row]
                continue

            record = {}
            row_len = len(raw_row)
            for idx, header in enumerate(headers):
                value = raw_row[idx].strip() if idx < row_len and raw_row[idx] is not None else ''
                record[header] = value
            rows.append(record)

    return rows


def build_room_value_lookup(rows):
    room_cols = [normalize_header(column_room_number_name)]
    for name in column_room_number_fallback_names:
        room_cols.append(normalize_header(name))

    value_col = normalize_header(column_heigth_name)
    fallback_col = normalize_header(column_ceiling_type_name)
    layer_col = 'layer'
    target_layer = normalize_key(excel_layer_column_value)

    def _build_lookup(apply_layer_filter):
        lookup = {}
        matched_rows = 0

        for row in rows:
            if apply_layer_filter and target_layer:
                row_layer = normalize_key(row.get(layer_col, ''))
                if row_layer != target_layer:
                    continue

            room_key = ''
            for room_col in room_cols:
                room_key = normalize_key(row.get(room_col, ''))
                if room_key:
                    break
            if not room_key:
                continue

            size_text = (row.get(value_col, '') or '').strip()
            ceiling_type = (row.get(fallback_col, '') or '').strip()
            value = build_ceiling_text(ceiling_type, size_text)

            matched_rows += 1
            if room_key not in lookup:
                lookup[room_key] = value

        return lookup, matched_rows

    lookup, matched_rows = _build_lookup(apply_layer_filter=True)
    if target_layer and matched_rows == 0:
        lookup, matched_rows = _build_lookup(apply_layer_filter=False)

    return lookup, matched_rows


def is_valid_element_id(element_id):
    if not element_id:
        return False
    try:
        return element_id != ElementId.InvalidElementId
    except Exception:
        return True


def read_room_identity(room):
    if not room:
        return "N/A", "N/A"

    try:
        name = getattr(room, 'Name', None) or "N/A"
    except Exception:
        name = "N/A"

    try:
        number = getattr(room, 'Number', None) or "N/A"
    except Exception:
        number = "N/A"

    return name, number


def get_room_from_link_reference(link_reference):
    """Resolve a linked room from LinkElementId-style references."""
    if not link_reference:
        return None

    try:
        link_instance_id = getattr(link_reference, 'LinkInstanceId', None)
        linked_element_id = getattr(link_reference, 'LinkedElementId', None)

        if is_valid_element_id(link_instance_id) and is_valid_element_id(linked_element_id):
            link_instance = doc.GetElement(link_instance_id)
            if link_instance and hasattr(link_instance, 'GetLinkDocument'):
                link_doc = link_instance.GetLinkDocument()
                if link_doc:
                    return link_doc.GetElement(linked_element_id)
    except Exception:
        pass

    return None


def resolve_tagged_room(tag):
    """Resolve room element for local and linked room tags across API versions."""
    # Path 1: direct Room property (host model rooms)
    try:
        if hasattr(tag, 'Room') and tag.Room:
            return tag.Room
    except Exception:
        pass

    # Path 2: local room id on room tag
    try:
        local_room_id = getattr(tag, 'TaggedLocalRoomId', None)
        if is_valid_element_id(local_room_id):
            room = doc.GetElement(local_room_id)
            if room:
                return room
    except Exception:
        pass

    # Path 3: tagged room id can be local ElementId or linked reference depending on API/version
    try:
        tagged_room_id = getattr(tag, 'TaggedRoomId', None)
        if tagged_room_id:
            if isinstance(tagged_room_id, ElementId):
                if is_valid_element_id(tagged_room_id):
                    room = doc.GetElement(tagged_room_id)
                    if room:
                        return room
            else:
                room = get_room_from_link_reference(tagged_room_id)
                if room:
                    return room
    except Exception:
        pass

    # Path 4: IndependentTag API local element ids
    try:
        for local_id in list(tag.GetTaggedLocalElementIds() or []):
            if is_valid_element_id(local_id):
                room = doc.GetElement(local_id)
                if room:
                    return room
    except Exception:
        pass

    # Path 5: IndependentTag API linked element ids
    try:
        for linked_ref in list(tag.GetTaggedElementIds() or []):
            room = get_room_from_link_reference(linked_ref)
            if room:
                return room
    except Exception:
        pass

    return None


def get_point_from_element(element):
    """Get a representative XYZ point for tags/notes."""
    if not element:
        return None

    for attr in ['TagHeadPosition', 'Coord']:
        try:
            pt = getattr(element, attr, None)
            if pt:
                return pt
        except Exception:
            pass

    try:
        location = getattr(element, 'Location', None)
        if location and hasattr(location, 'Point'):
            return location.Point
    except Exception:
        pass

    return None


def project_axis_value(point, axis, origin):
    """Project point on view axis for screen-relative comparisons."""
    if not point or not axis:
        return None

    dx = point.X - origin.X
    dy = point.Y - origin.Y
    dz = point.Z - origin.Z
    return (dx * axis.X) + (dy * axis.Y) + (dz * axis.Z)


def find_note_below_tag(tag, indexed_notes, used_note_ids, max_dx=12.0, max_dy=12.0):
    """Find nearest TextNote below tag in current view orientation."""
    tag_point = get_point_from_element(tag)
    if not tag_point:
        return None

    try:
        up_axis = view.UpDirection
        right_axis = view.RightDirection
        origin = view.Origin
    except Exception:
        return None

    tag_up = project_axis_value(tag_point, up_axis, origin)
    tag_right = project_axis_value(tag_point, right_axis, origin)
    if tag_up is None or tag_right is None:
        return None

    best_note = None
    best_score = None

    for note, note_up, note_right in indexed_notes:
        try:
            if note.Id in used_note_ids:
                continue
        except Exception:
            pass

        if note_up is None or note_right is None:
            continue

        # "Below" means lower on screen, i.e., smaller projection on UpDirection.
        dy = tag_up - note_up
        if dy <= 0 or dy > max_dy:
            continue

        dx = abs(tag_right - note_right)
        if dx > max_dx:
            continue

        # Prefer mostly vertical alignment, then nearest vertical distance.
        score = (dx * 2.0) + dy
        if best_score is None or score < best_score:
            best_score = score
            best_note = note

    return best_note


# Main Code
# ==================================================
try:
    csv_path = pick_csv_file()
    if not csv_path:
        output.print_md('## Cancelled: no CSV selected.')
        script.exit()

    csv_rows = load_csv_rows(csv_path)
    room_value_lookup, csv_match_count = build_room_value_lookup(csv_rows)

    room_tags = list((FilteredElementCollector(doc, view.Id)
                      .OfCategory(BuiltInCategory.OST_RoomTags)
                      .WhereElementIsNotElementType()
                      .ToElements()))

    if len(room_tags) == 0:
        output.print_md("## No room tags found in current view.")
        script.exit()

    tags = [t for t in room_tags if t.OwnerViewId == view.Id]
    if not tags:
        output.print_md("## No room tags found in current view.")
        script.exit()

    text_notes = list((FilteredElementCollector(doc, view.Id)
                       .OfClass(TextNote)
                       .WhereElementIsNotElementType()
                       .ToElements()))

    # Pre-index note positions once for fast nearest-note matching.
    indexed_notes = []
    used_note_ids = set()
    updates = []
    try:
        up_axis = view.UpDirection
        right_axis = view.RightDirection
        origin = view.Origin
        for note in text_notes:
            pt = get_point_from_element(note)
            note_up = project_axis_value(pt, up_axis, origin) if pt else None
            note_right = project_axis_value(pt, right_axis, origin) if pt else None
            indexed_notes.append((note, note_up, note_right))
    except Exception:
        indexed_notes = []

    output.print_md("# Room Tags in View")
    output.print_md("---")

    for i, tag in enumerate(tags, start=1):
        room_name = "N/A"
        room_number = "N/A"
        tag_id = tag.Id

        # Method 1: Try tag parameters directly (Room Name, Room Number)
        try:
            name_param = tag.LookupParameter("Room Name")
            if name_param and name_param.AsString():
                room_name = name_param.AsString()
        except Exception:
            pass

        try:
            number_param = tag.LookupParameter("Room Number")
            if number_param and number_param.AsString():
                room_number = number_param.AsString()
        except Exception:
            pass

        # Method 2: Resolve tagged room from host/linked model APIs
        try:
            if is_missing(room_name) or is_missing(room_number):
                room = resolve_tagged_room(tag)
                resolved_name, resolved_number = read_room_identity(room)
                if is_missing(room_name) and not is_missing(resolved_name):
                    room_name = resolved_name
                if is_missing(room_number) and not is_missing(resolved_number):
                    room_number = resolved_number
        except Exception:
            pass

        below_text = "N/A"
        desired_text = None

        try:
            matched_note = find_note_below_tag(tag, indexed_notes, used_note_ids)
            if matched_note:
                used_note_ids.add(matched_note.Id)
                raw_text = matched_note.Text if hasattr(matched_note, 'Text') else ""
                if raw_text:
                    below_text = raw_text.replace('\n', ' ').replace('\r', '').strip()
                room_keys = [normalize_key(room_number), normalize_key(room_name)]
                for room_key in room_keys:
                    if room_key and room_key in room_value_lookup:
                        desired_text = room_value_lookup[room_key]
                        break

                if not desired_text:
                    # Keep existing valid ACT/GWB/OTS note when no room/csv match is found.
                    desired_text = below_text if looks_like_valid_ceiling_text(below_text) else default_second_tag_text

                updates.append((tag, matched_note, desired_text))
        except Exception:
            pass

        if not desired_text:
            desired_text = default_second_tag_text

        output.print_md(
            '### {:03d}: ID {} | Room: {} | Number: {} | {}: {} | New: {}'.format(
                i,
                output.linkify(tag_id),
                room_name,
                room_number,
                column_heigth_name,
                below_text,
                desired_text,
            )
        )

    changed_notes = 0
    with revit.Transaction('Update room text notes from CSV'):
        for _, note, new_text in updates:
            try:
                old_text = note.Text if hasattr(note, 'Text') else ''
                if old_text != new_text:
                    note.Text = new_text
                    changed_notes += 1
            except Exception:
                pass

    output.print_md("---")
    output.print_md("**Total: {} room tags**".format(len(tags)))
    output.print_md("**CSV file: {}**".format(csv_path))
    output.print_md("**CSV matched rows: {}**".format(csv_match_count))
    output.print_md("**Text notes updated: {}**".format(changed_notes))

    if text_notes:
        output.print_md("---")
        output.print_md("# Text Notes in View")
        output.print_md("---")

        for i, note in enumerate(text_notes, start=1):
            note_text = note.Text if hasattr(note, 'Text') else "N/A"
            # Clean up newlines and extra whitespace
            if note_text and note_text != "N/A":
                note_text = note_text.replace('\n', ' ').replace('\r', '').strip()
            note_id = note.Id

            output.print_md(
                '### {:03d}: ID {} | Text: "{}"'.format(
                    i,
                    output.linkify(note_id),
                    note_text,
                )
            )

        output.print_md("---")
        output.print_md("**Total: {} text notes**".format(len(text_notes)))

except Exception as e:
    output.print_md("## Error: {}".format(str(e)))
    script.exit()
