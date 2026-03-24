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
    BuiltInParameter,
    ElementId,
    TextNote,
    TextNoteOptions,
    TextNoteType,
    HorizontalTextAlignment,
    XYZ,
)
from pyrevit import revit, script
from System.Collections.Generic import List
import math


# Button info
# ===================================================
__title__ = "Add Note Below Room Label"
__doc__ = """
Selects all Room Tags in the active view
"""

# Variables
# ==================================================
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
view = revit.active_view
TARGET_TEXT_TYPE_NAME = '1/8" Calibri 2 - Red'
DUPLICATE_DISTANCE_FT = 2.0  # Tolerance for detecting duplicates (in feet) - more permissive


def distance_between_points(pt1, pt2):
    """Calculate 3D distance between two XYZ points."""
    dx = pt1.X - pt2.X
    dy = pt1.Y - pt2.Y
    dz = pt1.Z - pt2.Z
    return math.sqrt(dx * dx + dy * dy + dz * dz)


def point_exists_nearby(existing_points, proposed_pt, tolerance_ft):
    """Check if a point already exists within tolerance distance of proposed location."""
    for existing_pt in existing_points:
        dist = distance_between_points(proposed_pt, existing_pt)
        if dist < tolerance_ft:
            return True
    return False


# Main Code
# ==================================================
try:
    room_tags = list((FilteredElementCollector(doc, view.Id)
                      .OfCategory(BuiltInCategory.OST_RoomTags)
                      .WhereElementIsNotElementType()
                      .ToElements()))

    if len(room_tags) == 0:
        script.exit()

    tags = [t for t in room_tags if t.OwnerViewId == view.Id]
    if not tags:
        script.exit()

    selected_ids = List[ElementId]([t.Id for t in tags])
    uidoc.Selection.SetElementIds(selected_ids)

    text_types = list(FilteredElementCollector(doc).OfClass(TextNoteType).ToElements())
    if not text_types:
        script.exit()

    text_type_id = None
    for txt_type in text_types:
        name_param = txt_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        type_name = name_param.AsString() if name_param else None
        if type_name == TARGET_TEXT_TYPE_NAME:
            text_type_id = txt_type.Id
            break

    if text_type_id is None:
        text_type_id = text_types[0].Id

    NOTE_TEXT = "XXX 09'-00\""
    OFFSET_FT = 1.0
    up = view.UpDirection

    # Collect existing text note positions - try both with and without view filter
    existing_text_notes_in_view = list((FilteredElementCollector(doc, view.Id)
                                        .OfClass(TextNote)
                                        .ToElements()))

    existing_text_notes_all = list((FilteredElementCollector(doc)
                                    .OfClass(TextNote)
                                    .ToElements()))

    print("Debug: TextNotes in view: {}, TextNotes in doc: {}".format(
        len(existing_text_notes_in_view), len(existing_text_notes_all)))

    # Use the full document collection since view-based might be filtering them out
    existing_text_notes = existing_text_notes_all
    existing_points = []
    error_count = 0
    for i, note in enumerate(existing_text_notes):
        try:
            note_loc = note.Location
            if note_loc:
                pt = note_loc.Origin
                if pt:
                    existing_points.append(pt)
                else:
                    if i < 5:
                        print("Note {} has Location but no Origin".format(i))
            else:
                if i < 5:
                    print("Note {} has no Location".format(i))
        except Exception as e:
            error_count += 1
            if error_count <= 5:
                print("Error getting note location: {}".format(str(e)))

    print("Extracted {} points from {} notes, {} errors".format(
        len(existing_points), len(existing_text_notes), error_count))

    with revit.Transaction("Add Room Tag Notes"):
        notes_created = 0
        notes_skipped = 0
        created_points = []
        debug_count = 0

        for tag in tags:
            loc = tag.Location
            if loc is None:
                continue
            pt = loc.Point
            note_pt = XYZ(
                pt.X - up.X * OFFSET_FT,
                pt.Y - up.Y * OFFSET_FT,
                pt.Z - up.Z * OFFSET_FT,
            )

            # Check against both existing notes AND newly created notes
            is_duplicate = False
            if point_exists_nearby(existing_points, note_pt, DUPLICATE_DISTANCE_FT):
                is_duplicate = True
            if point_exists_nearby(created_points, note_pt, DUPLICATE_DISTANCE_FT):
                is_duplicate = True

            if is_duplicate:
                notes_skipped += 1
                debug_count += 1
                if debug_count <= 10:  # Print first 10 skips for debugging
                    print("Skipping duplicate at ({}, {}, {})".format(note_pt.X, note_pt.Y, note_pt.Z))
                continue

            opts = TextNoteOptions(text_type_id)
            opts.HorizontalAlignment = HorizontalTextAlignment.Center
            new_note = TextNote.Create(doc, view.Id, note_pt, NOTE_TEXT, opts)
            if new_note:
                created_points.append(note_pt)
                notes_created += 1

        print("Created: {}, Skipped (duplicates): {}, Existing notes: {}".format(
            notes_created, notes_skipped, len(existing_points)))

except Exception as e:
    script.exit()
