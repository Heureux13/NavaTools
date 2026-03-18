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
    TextNoteOptions,
    TextNoteType,
    HorizontalTextAlignment,
    XYZ,
)
from pyrevit import revit, script
from System.Collections.Generic import List


# Button info
# ===================================================
__title__ = "Isolate by Rooms"
__doc__ = """
Selects all Room Tags in the active view
"""

# Variables
# ==================================================
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
view = revit.active_view


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

    text_types = list(FilteredElementCollector(doc).OfClass(TextNoteType).ToElementIds())
    if not text_types:
        script.exit()
    text_type_id = text_types[0]

    NOTE_TEXT = "XXX 09'-00\""
    OFFSET_FT = 0.5
    up = view.UpDirection

    with revit.Transaction("Add Room Tag Notes"):
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
            opts = TextNoteOptions(text_type_id)
            opts.HorizontalAlignment = HorizontalTextAlignment.Center
            TextNote.Create(doc, view.Id, note_pt, NOTE_TEXT, opts)

except Exception as e:
    script.exit()
