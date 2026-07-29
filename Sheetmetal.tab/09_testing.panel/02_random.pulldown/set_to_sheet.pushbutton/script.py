# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import forms, revit, script
from Autodesk.Revit.DB import (
    BuiltInCategory,
    FamilySymbol,
    FilteredElementCollector,
    View,
    Viewport,
    ViewSheet,
    XYZ,
)

# SheetCollection is a native Revit 2024+ feature
try:
    from Autodesk.Revit.DB import SheetCollection
    _SHEET_COLLECTION_API = True
except Exception:
    _SHEET_COLLECTION_API = False


# Button info
# ======================================================================
__title__ = 'Set to Sheet'
__doc__ = '''
Take view and center it on sheet
'''


# Variables
# ======================================================================
doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


class ListOption(object):
    def __init__(self, item, display_name):
        self.item = item
        self.display_name = display_name


def _safe_name(element):
    try:
        return element.Name.lower()
    except Exception:
        return ''


def get_view_display_name(view):
    try:
        return '{} | {}'.format(view.ViewType, view.Name)
    except Exception:
        return str(view.Id)


def get_titleblock_display_name(symbol):
    try:
        family = symbol.Family
        family_name = family.Name if family else "Title Block"
    except Exception:
        family_name = "Title Block"
    try:
        sym_name = symbol.Name
    except Exception:
        sym_name = str(symbol.Id)
    return "{} | {}".format(family_name, sym_name)


def get_unique_sheet_number(pattern, j, existing_numbers):
    while True:
        candidate = pattern.format(j)
        if candidate not in existing_numbers:
            existing_numbers.add(candidate)
            return candidate, j + 1
        j += 1


def collect_titleblocks(document):
    return sorted(
        FilteredElementCollector(document)
        .OfClass(FamilySymbol)
        .OfCategory(BuiltInCategory.OST_TitleBlocks)
        .ToElements(),
        key=lambda symbol: get_titleblock_display_name(symbol).lower()
    )


def collect_selected_views(document, active_uidoc):
    selected = []
    for element_id in active_uidoc.Selection.GetElementIds():
        element = document.GetElement(element_id)
        if (isinstance(element, View)
                and not element.IsTemplate
                and not isinstance(element, ViewSheet)):
            selected.append(element)
    return sorted(selected, key=lambda v: get_view_display_name(v).lower())


def collect_all_views(document):
    return sorted(
        [
            v for v in FilteredElementCollector(document).OfClass(View)
            if not v.IsTemplate and not isinstance(v, ViewSheet)
        ],
        key=lambda v: get_view_display_name(v).lower()
    )


def get_sheet_center(sheet):
    try:
        outline = sheet.Outline
        if outline is not None and (outline.Max.U - outline.Min.U) > 0.001:
            return XYZ(
                (outline.Min.U + outline.Max.U) / 2.0,
                (outline.Min.V + outline.Max.V) / 2.0,
                0
            )
    except Exception:
        pass
    # Fallback: safe centre for a typical 30"x42" D-size sheet (in feet)
    return XYZ(1.25, 0.875, 0)


_NO_COLLECTION = '(No Collection)'
_NEW_COLLECTION = '+ Create New...'


def collect_sheet_collections(document):
    if not _SHEET_COLLECTION_API:
        return []
    try:
        return sorted(
            FilteredElementCollector(document).OfClass(SheetCollection).ToElements(),
            key=_safe_name
        )
    except Exception:
        return []


def prompt_for_sheet_collection(document):
    """Return (existing_collection, new_name). Exactly one is set, or both None to skip."""
    if not _SHEET_COLLECTION_API:
        return None, None

    existing = collect_sheet_collections(document)
    options = [ListOption(None, _NO_COLLECTION)]
    for coll in existing:
        try:
            options.append(ListOption(coll, coll.Name))
        except Exception:
            pass
    options.append(ListOption(_NEW_COLLECTION, _NEW_COLLECTION))

    picked = forms.SelectFromList.show(
        options,
        name_attr='display_name',
        multiselect=False,
        title='Assign to Sheet Collection'
    )
    if not picked or picked.item is None:
        return None, None

    if picked.item == _NEW_COLLECTION:
        new_name = forms.ask_for_string(
            default='',
            prompt='Enter a name for the new sheet collection:',
            title='New Sheet Collection'
        )
        if not new_name or not new_name.strip():
            return None, None
        return None, new_name.strip()

    return picked.item, None


# Step 1: select views
# ======================================================================
selected_views = collect_selected_views(doc, uidoc)
if not selected_views:
    all_views = collect_all_views(doc)
    if not all_views:
        output.print_md('**No views found in this project.**')
        script.exit()

    view_options = [ListOption(v, get_view_display_name(v)) for v in all_views]
    picked_views = forms.SelectFromList.show(
        view_options,
        name_attr='display_name',
        multiselect=True,
        title='Select Views'
    )
    if not picked_views:
        script.exit()
    selected_views = [opt.item for opt in picked_views]

# Step 2: title block
# ======================================================================
titleblocks = collect_titleblocks(doc)
if not titleblocks:
    output.print_md('**No title blocks found in this project.**')
    script.exit()

titleblock_options = [
    ListOption(symbol, get_titleblock_display_name(symbol)) for symbol in titleblocks
]
selected_titleblock = forms.SelectFromList.show(
    titleblock_options,
    name_attr='display_name',
    multiselect=False,
    title='Select Title Block'
)
if not selected_titleblock:
    script.exit()

# Step 3: sheet number prefix
# ======================================================================
sheet_prefix = forms.ask_for_string(
    default='M-',
    prompt='Enter a sheet number prefix.\nSheets are numbered PREFIX001, PREFIX002, etc.',
    title='Sheet Number Prefix'
)
if sheet_prefix is None:
    script.exit()
sheet_prefix = sheet_prefix.strip()
if not sheet_prefix:
    output.print_md('**Sheet number prefix cannot be blank.**')
    script.exit()

# Step 4: sheet collection (Revit 2024+ only; silently skipped on older versions)
# ======================================================================
collection_element, collection_new_name = prompt_for_sheet_collection(doc)

# Existing sheet numbers to avoid collisions
existing_sheet_numbers = {
    s.SheetNumber for s in FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
    if getattr(s, 'SheetNumber', None)
}

created = []
failed = []
counter = 1

with revit.Transaction('Create Sheets'):
    for view in selected_views:
        sheet = None
        try:
            sheet = ViewSheet.Create(doc, selected_titleblock.item.Id)

            number_pattern = sheet_prefix + '{:03d}'
            sheet_number, counter = get_unique_sheet_number(
                number_pattern, counter, existing_sheet_numbers)
            sheet.SheetNumber = sheet_number
            sheet.Name = view.Name

            # Regenerate so sheet Outline and view geometry are valid
            doc.Regenerate()

            if not Viewport.CanAddViewToSheet(doc, sheet.Id, view.Id):
                raise Exception('view is already placed on another sheet')

            Viewport.Create(doc, sheet.Id, view.Id, get_sheet_center(sheet))
            created.append((sheet, view))
        except Exception as err:
            try:
                if sheet is not None:
                    doc.Delete(sheet.Id)
            except Exception:
                pass
            failed.append('{} ({})'.format(
                view.Name if hasattr(view, 'Name') else str(view.Id), err))

    # Assign to sheet collection (Revit 2024+)
    # Sheets are assigned to a collection via a parameter on the ViewSheet itself,
    # not by calling a method on the SheetCollection element.
    if created and _SHEET_COLLECTION_API and (collection_element is not None or collection_new_name):
        try:
            target_collection = collection_element
            if collection_new_name:
                target_collection = SheetCollection.Create(doc, collection_new_name)

            coll_id = target_collection.Id
            assign_ok = 0
            assign_fail = []
            for s, _ in created:
                param = s.LookupParameter('Sheet Collection')
                if param is None or param.IsReadOnly:
                    assign_fail.append(s.SheetNumber)
                else:
                    param.Set(coll_id)
                    assign_ok += 1

            if assign_ok:
                output.print_md('Assigned {} sheet(s) to collection **{}**.'.format(
                    assign_ok, target_collection.Name))
            if assign_fail:
                output.print_md(
                    '**Could not assign {} sheet(s) to collection** '
                    '(parameter missing or read-only): {}'.format(
                        len(assign_fail), ', '.join(assign_fail)))
        except Exception as coll_err:
            output.print_md('**Could not assign to sheet collection:** {}'.format(coll_err))

if created:
    output.print_md('# Created {} sheet(s)'.format(len(created)))
    for sheet, view in created:
        output.print_md('- {} -> {}'.format(view.Name, output.linkify(sheet.Id)))

if failed:
    output.print_md('---')
    output.print_md('## Could not create {} sheet(s)'.format(len(failed)))
    for message in failed:
        output.print_md('- {}'.format(message))
