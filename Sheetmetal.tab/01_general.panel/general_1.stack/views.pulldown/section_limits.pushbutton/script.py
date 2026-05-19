# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import revit, DB, forms, script

# Button info
# ======================================================================
__title__ = 'Section Limits'
__doc__ = '''Set the top and bottom elevation of selected Section views
relative to a chosen level with a custom offset above and below.

Usage:
1. Open a Section view OR select viewports on a sheet, then run this tool
2. Pick a reference level
3. Enter how far above the level for the TOP
4. Enter how far below the level for the BOTTOM
'''

# Variables
# ======================================================================

output = script.get_output()
doc = revit.doc
uidoc = revit.uidoc


# ── 1. Get selected section views ────────────────────────────────────
# Priority: selected viewports on a sheet → direct View selection → active view
sel_ids = uidoc.Selection.GetElementIds()
section_views = []

for eid in sel_ids:
    elem = doc.GetElement(eid)
    if isinstance(elem, DB.Viewport):
        view = doc.GetElement(elem.ViewId)
        if view and view.ViewType == DB.ViewType.Section:
            section_views.append(view)
    elif isinstance(elem, DB.View) and elem.ViewType == DB.ViewType.Section:
        section_views.append(elem)

# Fall back to the currently open/active view
if not section_views:
    active = revit.active_view
    if active and active.ViewType == DB.ViewType.Section:
        section_views.append(active)

if not section_views:
    output.print_md('## Section Limits')
    output.print_md(
        '**Error:** No Section view found.\n\n'
        'Open a Section view or select one or more section viewports on a sheet, then run this tool.'
    )
    script.exit()


def ft_to_ft_in(feet):
    total_inches = int(round(feet * 12))
    ft = total_inches // 12
    inch = total_inches % 12
    return "{}'  {}\"".format(ft, inch)


# ── 2. Which limit to set: Top or Bottom ─────────────────────────────
edge = forms.SelectFromList.show(
    ['Top', 'Bottom'],
    title='Which limit do you want to set?',
    multiselect=False
)
if not edge:
    script.exit()


# ── 3. Select reference level ─────────────────────────────────────────
all_levels = sorted(
    DB.FilteredElementCollector(doc).OfClass(DB.Level).ToElements(),
    key=lambda l: l.Elevation
)
level_names = ['{}   ({})'.format(l.Name, ft_to_ft_in(l.Elevation))
               for l in all_levels]

selected_key = forms.SelectFromList.show(
    level_names,
    title='Select Reference Level',
    multiselect=False,
    sort_keys=False
)
if not selected_key:
    script.exit()

ref_level = all_levels[level_names.index(selected_key)]
level_elev = ref_level.Elevation  # internal feet


# ── 4. Above or below that level ─────────────────────────────────────
direction = forms.SelectFromList.show(
    ['Above', 'Below'],
    title='Above or Below "' + ref_level.Name + '"?',
    multiselect=False
)
if not direction:
    script.exit()


# ── 5. How many inches ────────────────────────────────────────────────
inches_str = forms.ask_for_string(
    prompt='{} of section  –  {} "{}" by how many inches?'.format(
        edge, direction, ref_level.Name
    ),
    title='Offset (inches)',
    default='6'
)
if inches_str is None:
    script.exit()

try:
    offset_in = float(inches_str)
except ValueError:
    output.print_md('**Error:** Invalid value – enter a number (e.g. 6).')
    script.exit()

offset_ft = offset_in / 12.0
if direction == 'Above':
    target_z = level_elev + offset_ft
else:
    target_z = level_elev - offset_ft


# ── 6. Apply to each section view ────────────────────────────────────
updated = []
skipped = []

with revit.Transaction('Set Section {} Limit'.format(edge)):
    for view in section_views:
        crop_box = view.CropBox
        transform = crop_box.Transform
        basis_y_z = transform.BasisY.Z

        if abs(basis_y_z) < 1e-6:
            skipped.append(view.Name)
            continue

        origin_z = transform.Origin.Z
        new_local_y = (target_z - origin_z) / basis_y_z

        if edge == 'Top':
            # Guard: top must stay above current bottom
            if new_local_y <= crop_box.Min.Y:
                skipped.append(
                    '{} (new top would be at or below current bottom)'.format(view.Name))
                continue
            new_min = crop_box.Min
            new_max = DB.XYZ(crop_box.Max.X, new_local_y, crop_box.Max.Z)
        else:
            # Guard: bottom must stay below current top
            if new_local_y >= crop_box.Max.Y:
                skipped.append(
                    '{} (new bottom would be at or above current top)'.format(view.Name))
                continue
            new_min = DB.XYZ(crop_box.Min.X, new_local_y, crop_box.Min.Z)
            new_max = crop_box.Max

        new_bb = DB.BoundingBoxXYZ()
        new_bb.Transform = transform
        new_bb.Min = new_min
        new_bb.Max = new_max
        view.CropBox = new_bb

        if not view.CropBoxActive:
            view.CropBoxActive = True

        updated.append(view.Name)


# ── 7. Report ─────────────────────────────────────────────────────────
output.print_md('## Section Limits Applied')
output.print_md('**Edge:** {}'.format(edge))
output.print_md(
    '**Level:** {}  @ {}'.format(ref_level.Name, ft_to_ft_in(level_elev)))
output.print_md('**Direction:** {} by {} in  →  {}'.format(direction,
                offset_in, ft_to_ft_in(target_z)))
output.print_md('---')

if updated:
    output.print_md('**Updated ({}):**'.format(len(updated)))
    for name in updated:
        output.print_md('- ' + name)

if skipped:
    output.print_md('**Skipped ({}):**'.format(len(skipped)))
    for name in skipped:
        output.print_md('- ' + name)
