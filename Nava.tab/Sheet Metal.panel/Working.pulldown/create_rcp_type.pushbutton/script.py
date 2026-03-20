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
DUPLICATE_DISTANCE_FT = 0.5  # Tolerance for detecting duplicates


def distance_between_points(pt1, pt2):
    """Calculate 3D distance between two XYZ points."""
    dx = pt1.X - pt2.X
    dy = pt1.Y - pt2.Y
    dz = pt1.Z - pt2.Z
    return math.sqrt(dx * dx + dy * dy + dz * dz)


def note_exists_nearby(existing_notes, proposed_pt, tolerance_ft):
    """Check if a text note already exists within tolerance distance of proposed location."""
    for note in existing_notes:
        try:
            note_loc = note.Location
            if note_loc:
                note_pt = note_loc.Point
                dist = distance_between_points(proposed_pt, note_pt)
                if dist < tolerance_ft:
                    return True
        except Exception:
            pass
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

    # Collect existing text notes to avoid duplicates
    existing_text_notes = list((FilteredElementCollector(doc, view.Id)
                                .OfClass(TextNote)
                                .ToElements()))

    with revit.Transaction("Add Room Tag Notes"):
        notes_created = 0
        notes_skipped = 0

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

            # Check if a note already exists at this location
            if note_exists_nearby(existing_text_notes, note_pt, DUPLICATE_DISTANCE_FT):
                notes_skipped += 1
                continue

            opts = TextNoteOptions(text_type_id)
            opts.HorizontalAlignment = HorizontalTextAlignment.Center
            new_note = TextNote.Create(doc, view.Id, note_pt, NOTE_TEXT, opts)
            if new_note:
                existing_text_notes.append(new_note)
                notes_created += 1

        print("Created: {}, Skipped (duplicates): {}".format(notes_created, notes_skipped))

except Exception as e:
    script.exit()
