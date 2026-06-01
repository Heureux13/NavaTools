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
# Prefer project elevation so behavior matches what users read in-project.
level_elev = ref_level.Elevation
try:
    level_elev = ref_level.ProjectElevation
except Exception:
    pass


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
        # Revit versions can expose section vertical on different local axes.
        axis_z_values = [
            transform.BasisX.Z,
            transform.BasisY.Z,
            transform.BasisZ.Z
        ]
        vert_axis_idx = max(range(3), key=lambda i: abs(axis_z_values[i]))
        vert_axis_z = axis_z_values[vert_axis_idx]

        if abs(vert_axis_z) < 1e-6:
            skipped.append(view.Name)
            continue

        min_vals = [crop_box.Min.X, crop_box.Min.Y, crop_box.Min.Z]
        max_vals = [crop_box.Max.X, crop_box.Max.Y, crop_box.Max.Z]

        if vert_axis_z > 0:
            top_bound = 'max'
            bottom_bound = 'min'
        else:
            top_bound = 'min'
            bottom_bound = 'max'

        center_vals = [
            (min_vals[0] + max_vals[0]) * 0.5,
            (min_vals[1] + max_vals[1]) * 0.5,
            (min_vals[2] + max_vals[2]) * 0.5
        ]

        if top_bound == 'max':
            top_local = max_vals[vert_axis_idx]
            bottom_local = min_vals[vert_axis_idx]
        else:
            top_local = min_vals[vert_axis_idx]
            bottom_local = max_vals[vert_axis_idx]

        top_pt_vals = list(center_vals)
        bottom_pt_vals = list(center_vals)
        top_pt_vals[vert_axis_idx] = top_local
        bottom_pt_vals[vert_axis_idx] = bottom_local
        top_pt = DB.XYZ(top_pt_vals[0], top_pt_vals[1], top_pt_vals[2])
        bottom_pt = DB.XYZ(bottom_pt_vals[0], bottom_pt_vals[1], bottom_pt_vals[2])

        top_world_z = transform.OfPoint(top_pt).Z
        bottom_world_z = transform.OfPoint(bottom_pt).Z

        # Solve local value from absolute world Z while keeping other axes fixed at center.
        axis_z_basis = [transform.BasisX.Z, transform.BasisY.Z, transform.BasisZ.Z]
        fixed_world_z = transform.Origin.Z
        for i in range(3):
            if i != vert_axis_idx:
                fixed_world_z += axis_z_basis[i] * center_vals[i]

        new_edge_local = (target_z - fixed_world_z) / vert_axis_z

        if edge == 'Top':
            new_top_local = new_edge_local

            # Guard in world coordinates to avoid local-axis ambiguity.
            if target_z <= bottom_world_z:
                skipped.append(
                    '{} (new top {} would be at or below current bottom {})'.format(
                        view.Name,
                        ft_to_ft_in(target_z),
                        ft_to_ft_in(bottom_world_z)
                    )
                )
                continue

            if top_bound == 'max':
                max_vals[vert_axis_idx] = new_top_local
            else:
                min_vals[vert_axis_idx] = new_top_local
        else:
            new_bottom_local = new_edge_local

            # Guard in world coordinates to avoid local-axis ambiguity.
            if target_z >= top_world_z:
                skipped.append(
                    '{} (new bottom {} would be at or above current top {})'.format(
                        view.Name,
                        ft_to_ft_in(target_z),
                        ft_to_ft_in(top_world_z)
                    )
                )
                continue

            if bottom_bound == 'min':
                min_vals[vert_axis_idx] = new_bottom_local
            else:
                max_vals[vert_axis_idx] = new_bottom_local

        crop_box.Min = DB.XYZ(min_vals[0], min_vals[1], min_vals[2])
        crop_box.Max = DB.XYZ(max_vals[0], max_vals[1], max_vals[2])
        view.CropBox = crop_box

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
