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
from System.Collections.Generic import List


# Button info
# ===================================================
__title__ = "List of room tags"
__doc__ = """
Lists all room tags in the active view.
"""

# Variables
# ==================================================
doc = revit.doc
view = revit.active_view
output = script.get_output()


# Main Code
# ==================================================
try:
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

        # Method 2: Try tag.Room property as fallback
        try:
            if room_name == "N/A" and hasattr(tag, 'Room') and tag.Room:
                room = tag.Room
                room_name = room.Name if hasattr(room, 'Name') else "N/A"
                room_number = room.Number if hasattr(room, 'Number') else "N/A"
        except Exception:
            pass

        # Method 3: Try tagged element IDs as last resort
        try:
            if room_name == "N/A":
                tagged_ids = list(tag.GetTaggedLocalElementIds() or [])
                if tagged_ids:
                    room_elem = doc.GetElement(tagged_ids[0])
                    if room_elem:
                        room_name = getattr(room_elem, 'Name', 'N/A')
                        room_number = getattr(room_elem, 'Number', 'N/A')
        except Exception:
            pass

        output.print_md(
            '### {:03d}: ID {} | Room: {} | Number: {}'.format(
                i,
                output.linkify(tag_id),
                room_name,
                room_number,
            )
        )

    output.print_md("---")
    output.print_md("**Total: {} room tags**".format(len(tags)))

    # Collect and display TextNotes
    text_notes = list((FilteredElementCollector(doc, view.Id)
                       .OfClass(TextNote)
                       .ToElements()))

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
