# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import forms, revit, script
from Autodesk.Revit.DB import BuiltInCategory
from Autodesk.Revit.DB import FilteredElementCollector
from Autodesk.Revit.DB import ViewSheet


__title__ = 'Duplicate Empty Sheet'
__doc__ = 'Select sheets and create empty copies.'


doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


class ListOption(object):
    def __init__(self, item, display_name):
        self.item = item
        self.display_name = display_name


def get_sheet_display_name(sheet):
    try:
        return '{} | {}'.format(sheet.SheetNumber, sheet.Name)
    except Exception:
        return str(sheet.Id)


def collect_selected_sheets(document, active_uidoc):
    selected = []
    ids = active_uidoc.Selection.GetElementIds()
    for element_id in ids:
        element = document.GetElement(element_id)
        if isinstance(element, ViewSheet):
            selected.append(element)
    selected.sort(key=lambda s: get_sheet_display_name(s).lower())
    return selected


def collect_all_sheets(document):
    sheets = list(FilteredElementCollector(document).OfClass(ViewSheet).ToElements())
    sheets.sort(key=lambda s: get_sheet_display_name(s).lower())
    return sheets


def get_sheet_titleblock_symbol_id(document, sheet):
    titleblocks = list(
        FilteredElementCollector(document, sheet.Id)
        .OfCategory(BuiltInCategory.OST_TitleBlocks)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    if not titleblocks:
        raise Exception('source sheet has no title block')
    return titleblocks[0].GetTypeId()


def get_unique_sheet_number(base_number, existing_numbers):
    if base_number not in existing_numbers:
        existing_numbers.add(base_number)
        return base_number

    index = 2
    while True:
        candidate = '{}-{:02d}'.format(base_number, index)
        if candidate not in existing_numbers:
            existing_numbers.add(candidate)
            return candidate
        index += 1


selected_sheets = collect_selected_sheets(doc, uidoc)
if not selected_sheets:
    all_sheets = collect_all_sheets(doc)
    if not all_sheets:
        output.print_md('**No sheets found in this project.**')
        script.exit()

    options = []
    for sheet in all_sheets:
        options.append(ListOption(sheet, get_sheet_display_name(sheet)))

    picked = forms.SelectFromList.show(
        options,
        name_attr='display_name',
        multiselect=True,
        title='Select Sheets to Duplicate Empty'
    )
    if not picked:
        script.exit()

    selected_sheets = []
    for option in picked:
        selected_sheets.append(option.item)


existing_sheet_numbers = set()
all_existing_sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
for sheet in all_existing_sheets:
    number = getattr(sheet, 'SheetNumber', None)
    if number:
        existing_sheet_numbers.add(number)


created = []
failed = []

with revit.Transaction('Duplicate Empty Sheets'):
    for source_sheet in selected_sheets:
        new_sheet = None
        try:
            titleblock_symbol_id = get_sheet_titleblock_symbol_id(doc, source_sheet)
            new_sheet = ViewSheet.Create(doc, titleblock_symbol_id)

            new_number = get_unique_sheet_number(source_sheet.SheetNumber, existing_sheet_numbers)
            new_sheet.SheetNumber = new_number
            new_sheet.Name = source_sheet.Name

            created.append((source_sheet, new_sheet))
        except Exception as err:
            try:
                if new_sheet is not None:
                    doc.Delete(new_sheet.Id)
            except Exception:
                pass
            failed.append('{} ({})'.format(get_sheet_display_name(source_sheet), err))


if created:
    output.print_md('# Created {} empty sheet(s)'.format(len(created)))
    for source_sheet, new_sheet in created:
        text = '- {} -> {}'.format(get_sheet_display_name(source_sheet), output.linkify(new_sheet.Id))
        output.print_md(text)

if failed:
    output.print_md('---')
    output.print_md('## Could not create {} sheet(s)'.format(len(failed)))
    for message in failed:
        output.print_md('- {}'.format(message))
