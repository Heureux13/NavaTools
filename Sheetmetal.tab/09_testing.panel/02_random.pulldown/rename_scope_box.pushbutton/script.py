# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import script, DB, forms, revit

# Button info
# ======================================================================
__title__ = 'Scope Box Rename'
__doc__ = '''
Select scope boxes and rename based on lettering: Area [?]01
'''

# Variables
# ======================================================================

output = script.get_output()

try:
    _revit_app = __revit__  # type: ignore[name-defined]
except NameError:
    _revit_app = None

uidoc = _revit_app.ActiveUIDocument if _revit_app else None
if uidoc is None:
    output.print_md('## Scope Box Rename')
    output.print_md('This command must run inside Revit.')
    raise RuntimeError('No ActiveUIDocument. Run inside Revit.')

doc = uidoc.Document


def get_selected_scope_boxes(document, selection_ids):
    boxes = []
    for eid in selection_ids:
        elem = document.GetElement(eid)
        if not elem or not elem.Category:
            continue
        if elem.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_VolumeOfInterest):
            boxes.append(elem)
    return boxes


def get_scope_box_center(scope_box):
    bbox = scope_box.get_BoundingBox(None)
    if not bbox:
        return DB.XYZ(0.0, 0.0, 0.0)

    transform = bbox.Transform if bbox.Transform else DB.Transform.Identity
    min_pt = bbox.Min
    max_pt = bbox.Max
    center_local = DB.XYZ(
        (min_pt.X + max_pt.X) / 2.0,
        (min_pt.Y + max_pt.Y) / 2.0,
        (min_pt.Z + max_pt.Z) / 2.0,
    )
    return transform.OfPoint(center_local)


def sorted_top_left(scope_boxes):
    # Top-left order in model coordinates: highest Y first, then lowest X.
    return sorted(scope_boxes, key=lambda sb: (-get_scope_box_center(sb).Y, get_scope_box_center(sb).X))


selection_ids = uidoc.Selection.GetElementIds()
scope_boxes = get_selected_scope_boxes(doc, selection_ids)

if not scope_boxes:
    forms.alert(
        'Select one or more Scope Boxes before running this command.',
        title='Scope Box Rename',
        exitscript=True
    )

affix = forms.ask_for_string(
    prompt='Enter affix letter/text for naming (example: A):',
    title='Scope Box Rename',
    default='A'
)

if affix is None:
    script.exit()

affix = affix.strip()
if not affix:
    forms.alert('Affix cannot be empty.', title='Scope Box Rename', exitscript=True)

ordered_boxes = sorted_top_left(scope_boxes)

all_scope_boxes = list(
    DB.FilteredElementCollector(doc)
    .OfCategory(DB.BuiltInCategory.OST_VolumeOfInterest)
    .WhereElementIsNotElementType()
    .ToElements()
)

selected_ids_int = set(sb.Id.IntegerValue for sb in ordered_boxes)
existing_names = set(
    sb.Name for sb in all_scope_boxes
    if sb.Id.IntegerValue not in selected_ids_int
)

number_width = max(2, len(str(len(ordered_boxes))))
target_names = [
    'Area {}{}'.format(affix, str(idx).zfill(number_width))
    for idx in range(1, len(ordered_boxes) + 1)
]
conflicts = [name for name in target_names if name in existing_names]

if conflicts:
    output.print_md('## Scope Box Rename')
    output.print_md('Rename stopped. The following names already exist on other scope boxes:')
    for name in sorted(set(conflicts)):
        output.print_md('- {}'.format(name))
    script.exit()

with revit.Transaction('Rename Scope Boxes'):
    # Temporary names avoid collision while selected scope boxes swap/shift names.
    for i, scope_box in enumerate(ordered_boxes):
        scope_box.Name = '__tmp_scope_rename_{}_{}'.format(scope_box.Id.IntegerValue, i)

    for index, scope_box in enumerate(ordered_boxes, start=1):
        scope_box.Name = 'Area {}{}'.format(affix, str(index).zfill(number_width))

output.print_md('## Scope Box Rename')
output.print_md('Renamed **{}** scope boxes using affix **{}**.'.format(len(ordered_boxes), affix))
