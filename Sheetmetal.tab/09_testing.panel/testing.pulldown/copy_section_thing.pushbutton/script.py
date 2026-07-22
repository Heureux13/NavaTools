# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.DB import (
    CopyPasteOptions,
    ElementId,
    ElementTransformUtils,
    Transform,
    Transaction,
    View,
    ViewSection,
    ViewType,
)
from pyrevit import forms, revit, script
from System.Collections.Generic import List

# Button info
# ======================================================================
__title__ = 'Copy Section'
__doc__ = '''
Use selected section view and copy it by a chosen spacing and direction.
Prompts for copy count (max 30), spacing, and direction.
'''

# Variables
# ======================================================================

output = script.get_output()


def _get_copy_count(max_count):
    user_input = forms.ask_for_string(
        default='1',
        prompt='How many copies do you want to make? (1-{})'.format(max_count),
        title='Copy Section'
    )
    if user_input is None:
        return None

    try:
        count = int(user_input.strip())
    except Exception:
        output.print_md(
            'Please enter a whole number between 1 and {}.'.format(max_count))
        return None

    if count < 1 or count > max_count:
        output.print_md(
            'Copy count must be between 1 and {}.'.format(max_count))
        return None

    return count


def _is_section_view(element):
    if element is not None and element.Category is not None:
        if element.Category.Name == 'Views':
            return True
    if isinstance(element, ViewSection):
        return True
    if isinstance(element, View):
        return element.ViewType == ViewType.Section
    return False


def _get_spacing_feet():
    user_input = forms.ask_for_string(
        default='8',
        prompt='Enter spacing in feet (example: 8 or 2.5):',
        title='Copy Section'
    )
    if user_input is None:
        return None

    try:
        spacing = float(user_input.strip())
    except Exception:
        output.print_md('Spacing must be a number in feet.')
        return None

    if spacing <= 0:
        output.print_md('Spacing must be greater than 0.')
        return None

    return spacing


def _get_direction_vector(active_view):
    direction_choice = forms.SelectFromList.show(
        ['Left', 'Right', 'Up', 'Down'],
        title='Copy Section - Direction',
        button_name='Use Direction',
        multiselect=False,
    )

    if not direction_choice:
        return None, None

    if direction_choice == 'Left':
        return direction_choice, -active_view.RightDirection
    if direction_choice == 'Right':
        return direction_choice, active_view.RightDirection
    if direction_choice == 'Up':
        return direction_choice, active_view.UpDirection
    return direction_choice, -active_view.UpDirection


def main():
    doc = revit.doc  # type: ignore[attr-defined]
    uidoc = getattr(revit, 'uidoc', None)
    if uidoc is None:
        revit_host = globals().get('__revit__')
        uidoc = revit_host.ActiveUIDocument if revit_host else None
    if uidoc is None:
        output.print_md('No active Revit UI document found.')
        return

    active_view = doc.ActiveView
    max_count = 30

    count = _get_copy_count(max_count)
    if count is None:
        output.print_md('Copy operation canceled.')
        return

    spacing_feet = _get_spacing_feet()
    if spacing_feet is None:
        output.print_md('Copy operation canceled.')
        return

    direction_name, move_direction = _get_direction_vector(active_view)
    if move_direction is None:
        output.print_md('Copy operation canceled.')
        return

    selected_ids = list(uidoc.Selection.GetElementIds())
    if not selected_ids:
        output.print_md(
            'Select the section view line (head+tail) first, then run the button.')
        return

    selected_sections = []
    selected_info = []
    for element_id in selected_ids:
        element = doc.GetElement(element_id)
        if _is_section_view(element):
            selected_sections.append(element_id)
        elem_type = type(
            element).__name__ if element is not None else 'NoneType'
        elem_cat = element.Category.Name if (
            element is not None and element.Category is not None) else 'No Category'
        selected_info.append('ID {} | Type {} | Category {}'.format(
            element_id.IntegerValue, elem_type, elem_cat))

    if not selected_sections:
        output.print_md('Current selection does not include a section view.')
        output.print_md('Selected elements detected:')
        for info_line in selected_info:
            output.print_md('- {}'.format(info_line))
        return

    if len(selected_sections) > 1:
        output.print_md('Select only one section view and run again.')
        return

    source_id = selected_sections[0]

    created_total = 0
    t = Transaction(doc, 'Copy Section')  # type: ignore[call-arg]
    t.Start()
    try:
        source_ids = List[ElementId]()  # type: ignore[type-arg]
        source_ids.Add(source_id)  # type: ignore[union-attr]
        copy_options = CopyPasteOptions()

        for i in range(1, count + 1):
            translation = move_direction.Multiply(spacing_feet * i)
            xform = Transform.CreateTranslation(  # type: ignore[arg-type]
                translation)  # type: ignore[arg-type]

            # Section heads/tails are view-specific symbols; use view-to-view copy.
            new_ids = ElementTransformUtils.CopyElements(  # type: ignore[arg-type]
                active_view,
                source_ids,
                active_view,
                xform,
                copy_options,
            )
            if new_ids and len(new_ids) > 0:
                created_total += 1
        t.Commit()
    except Exception as ex:
        if t.HasStarted():
            t.RollBack()
        output.print_md('Copy failed: {}'.format(ex))
        return

    output.print_md(
        'Created {} section copy/copies at {} ft spacing toward {}.'.format(
            created_total, spacing_feet, direction_name))


if __name__ == '__main__':
    main()
