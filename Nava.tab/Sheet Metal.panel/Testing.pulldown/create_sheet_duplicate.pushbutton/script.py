# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import forms, revit, script
import re
from System.Collections.Generic import List
from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    CopyPasteOptions,
    ElementTransformUtils,
    FilteredElementCollector,
    StorageType,
    Transform,
    View,
    Viewport,
    ViewSheet,
    ElementId,
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
__title__ = 'Create Duplicate Sheet'
__doc__ = '''
Select target views and reference sheets, then create new sheets with view locations matched by scope box.
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


def get_sheet_display_name(sheet):
    try:
        return '{} | {}'.format(sheet.SheetNumber, sheet.Name)
    except Exception:
        return str(sheet.Id)


def get_unique_scope_sheet_number(prefix, scope_suffix, existing_numbers):
    base = '{}{}'.format(prefix, scope_suffix)
    if base not in existing_numbers:
        existing_numbers.add(base)
        return base

    i = 2
    while True:
        candidate = '{}{:02d}-{}'.format(prefix, i, scope_suffix)
        if candidate not in existing_numbers:
            existing_numbers.add(candidate)
            return candidate
        i += 1


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


def get_sheet_titleblock_instance(document, sheet):
    tbs = list(
        FilteredElementCollector(document, sheet.Id)
        .OfCategory(BuiltInCategory.OST_TitleBlocks)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    if not tbs:
        return None
    return tbs[0]


def copy_writable_parameters(source_element, target_element, skip_param_names=None):
    if source_element is None or target_element is None:
        return

    skip_names = set()
    if skip_param_names:
        for name in skip_param_names:
            try:
                skip_names.add((name or '').strip().lower())
            except Exception:
                pass

    for src_param in list(source_element.Parameters):
        try:
            if src_param is None or not src_param.HasValue:
                continue
            definition = src_param.Definition
            if definition is None:
                continue
            dname = (definition.Name or '').strip().lower()
            if dname in skip_names:
                continue

            dst_param = target_element.LookupParameter(definition.Name)
            if dst_param is None or dst_param.IsReadOnly:
                continue
            if dst_param.StorageType != src_param.StorageType:
                continue

            stype = src_param.StorageType
            if stype == StorageType.Double:
                dst_param.Set(src_param.AsDouble())
            elif stype == StorageType.Integer:
                dst_param.Set(src_param.AsInteger())
            elif stype == StorageType.String:
                dst_param.Set(src_param.AsString())
            elif stype == StorageType.ElementId:
                dst_param.Set(src_param.AsElementId())
        except Exception:
            continue


def copy_titleblock_instance_parameters(source_tb, target_tb):
    copy_writable_parameters(source_tb, target_tb)


def collect_selected_views(document, active_uidoc):
    selected = []
    for element_id in active_uidoc.Selection.GetElementIds():
        element = document.GetElement(element_id)
        if (isinstance(element, View)
                and not element.IsTemplate
                and not isinstance(element, ViewSheet)):
            selected.append(element)
    return sorted(selected, key=lambda v: get_view_display_name(v).lower())


def collect_selected_sheets(document, active_uidoc):
    selected = []
    for element_id in active_uidoc.Selection.GetElementIds():
        element = document.GetElement(element_id)
        if isinstance(element, ViewSheet):
            selected.append(element)
    return sorted(selected, key=lambda s: get_sheet_display_name(s).lower())


def collect_all_sheets(document):
    return sorted(
        FilteredElementCollector(document).OfClass(ViewSheet).ToElements(),
        key=lambda s: get_sheet_display_name(s).lower()
    )


def collect_all_views(document):
    return sorted(
        [
            v for v in FilteredElementCollector(document).OfClass(View)
            if not v.IsTemplate and not isinstance(v, ViewSheet)
        ],
        key=lambda v: get_view_display_name(v).lower()
    )


def collect_view_templates(document):
    return sorted(
        [
            v for v in FilteredElementCollector(document).OfClass(View)
            if v.IsTemplate
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


def copy_reference_sheet_content(document, source_sheet, target_sheet):
    """Copy non-viewport, non-titleblock elements from source sheet to target sheet."""
    source_elements = list(
        FilteredElementCollector(document, source_sheet.Id)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    ids_to_copy = []
    for element in source_elements:
        if element is None:
            continue
        if isinstance(element, Viewport):
            continue
        try:
            category = element.Category
            if category and category.Id.IntegerValue == int(BuiltInCategory.OST_TitleBlocks):
                continue
        except Exception:
            pass
        ids_to_copy.append(element.Id)

    if not ids_to_copy:
        return

    ids_collection = List[ElementId]()
    for element_id in ids_to_copy:
        ids_collection.Add(element_id)

    copy_options = CopyPasteOptions()
    ElementTransformUtils.CopyElements(
        source_sheet,
        ids_collection,
        target_sheet,
        Transform.Identity,
        copy_options
    )


def get_viewport_center(vp):
    try:
        return vp.GetBoxCenter()
    except Exception:
        return None


def apply_viewport_layout(source_vp, target_vp):
    """Match viewport layout properties so placement mirrors the reference sheet."""
    if source_vp is None or target_vp is None:
        return

    try:
        target_vp.ChangeTypeId(source_vp.GetTypeId())
    except Exception:
        pass

    try:
        target_vp.Rotation = source_vp.Rotation
    except Exception:
        pass

    try:
        center = source_vp.GetBoxCenter()
        target_vp.SetBoxCenter(center)
    except Exception:
        pass


def get_scope_box_key(view):
    scope_id = None
    try:
        param = view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)
        if param:
            eid = param.AsElementId()
            if eid and eid != ElementId.InvalidElementId:
                scope_id = eid.IntegerValue
    except Exception:
        pass

    if scope_id is None:
        try:
            param = view.LookupParameter('Scope Box')
            if param:
                eid = param.AsElementId()
                if eid and eid != ElementId.InvalidElementId:
                    scope_id = eid.IntegerValue
        except Exception:
            pass

    if scope_id is None:
        return None
    return str(scope_id)


def get_scope_box_suffix(document, view):
    try:
        param = view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)
        if not param:
            return None
        scope_id = param.AsElementId()
        if not scope_id or scope_id == ElementId.InvalidElementId:
            return None

        scope = document.GetElement(scope_id)
        if scope is None:
            return None

        raw_name = (scope.Name or '').strip().upper()
        if not raw_name:
            return None

        # Example: "AREA F4" -> "F4"
        tokens = raw_name.replace('-', ' ').split()
        if not tokens:
            return None
        suffix = ''.join(ch for ch in tokens[-1] if ch.isalnum())
        if not suffix:
            return None
        return suffix
    except Exception:
        return None


def _set_param_value(element, param_name, value):
    if element is None:
        return False
    try:
        param = element.LookupParameter(param_name)
        if param is None or param.IsReadOnly:
            return False

        if param.StorageType == StorageType.String:
            param.Set(str(value))
            return True

        if param.StorageType == StorageType.Integer:
            if isinstance(value, bool):
                param.Set(1 if value else 0)
                return True
            if isinstance(value, int):
                param.Set(value)
                return True
            return False

        if param.StorageType == StorageType.Double and isinstance(value, (int, float)):
            param.Set(float(value))
            return True

        return False
    except Exception:
        return False


def _set_matching_area_flag(element, area_code):
    if element is None:
        return

    normalized_target = area_code.upper().replace('-', '')
    area_flag_names = []

    for param in list(element.Parameters):
        try:
            if param is None or param.Definition is None:
                continue
            pname = param.Definition.Name or ''
            if not pname.startswith('Area '):
                continue
            if pname in ('Area', 'Area #'):
                continue

            candidate_suffix = pname[5:].strip().upper().replace('-', '')
            if not candidate_suffix:
                continue

            area_flag_names.append((pname, candidate_suffix))
        except Exception:
            continue

    # Reset all area flags first so only one area stays checked.
    for pname, _ in area_flag_names:
        _set_param_value(element, pname, False)

    # Primary match: exact area code (e.g. Area C)
    for pname, candidate_suffix in area_flag_names:
        if candidate_suffix == normalized_target:
            _set_param_value(element, pname, True)
            return

    # Fallback: first flag starting with area code (e.g. Area C1)
    for pname, candidate_suffix in area_flag_names:
        if candidate_suffix.startswith(normalized_target):
            _set_param_value(element, pname, True)
            return


def split_scope_suffix(scope_suffix):
    cleaned = ''.join(ch for ch in (scope_suffix or '').upper() if ch.isalnum())
    area_code = ''.join(ch for ch in cleaned if ch.isalpha())
    area_number = ''.join(ch for ch in cleaned if ch.isdigit())

    if not area_code:
        area_code = cleaned

    return area_code, area_number


def split_area_from_sheet_name(sheet_name):
    # Expected suffix format: "... Area C2" or "... Area C 2"
    raw = (sheet_name or '').upper().strip()
    match = re.search(r'AREA\s+([A-Z]+)\s*([0-9]+)\s*$', raw)
    if not match:
        return None, None
    return match.group(1), match.group(2)


def apply_area_parameters(sheet, titleblock_instance, scope_suffix, source_sheet_name):
    area_code, area_number = split_area_from_sheet_name(source_sheet_name)
    if not area_code:
        area_code, area_number = split_scope_suffix(scope_suffix)

    for element in (sheet, titleblock_instance):
        _set_param_value(element, 'Area', area_code)
        _set_param_value(element, 'Area #', area_number)
        _set_param_value(element, 'Ref Sheet', '0')
        _set_param_value(element, 'Ref Sheet Trade', 'M')
        _set_param_value(element, 'Trade', 'MD')
        _set_matching_area_flag(element, area_code)


def build_reference_viewport_index(document, sheets):
    """
    Build an index: scope_box_key -> list[(sheet, viewport, view)].
    """
    index = {}
    for sheet in sheets:
        vps = list(
            FilteredElementCollector(document, sheet.Id)
            .OfClass(Viewport)
            .ToElements()
        )
        for vp in vps:
            src_view = document.GetElement(vp.ViewId)
            if src_view is None:
                continue
            key = get_scope_box_key(src_view)
            if not key:
                continue
            index.setdefault(key, []).append((sheet, vp, src_view))
    return index


def choose_best_reference(target_view, candidates, used_viewport_ids):
    if not candidates:
        return None

    best_candidate = None
    best_score = None
    for candidate in candidates:
        sheet, vp, src_view = candidate
        if vp.Id.IntegerValue in used_viewport_ids:
            continue

        score = 0
        if str(src_view.ViewType) == str(target_view.ViewType):
            score += 10
        try:
            if src_view.Scale == target_view.Scale:
                score += 3
        except Exception:
            pass

        if best_score is None or score > best_score:
            best_score = score
            best_candidate = candidate

    return best_candidate


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


def prompt_for_view_template(document):
    options = [ListOption(None, '(Keep Source View Template)')]
    for template in collect_view_templates(document):
        options.append(ListOption(template, get_view_display_name(template)))

    picked = forms.SelectFromList.show(
        options,
        name_attr='display_name',
        multiselect=False,
        title='Select View Template For Duplicated Views'
    )
    if not picked or picked.item is None:
        return ElementId.InvalidElementId
    return picked.item.Id


# Step 1: select target views
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
        title='Select Target Views'
    )
    if not picked_views:
        script.exit()
    selected_views = [opt.item for opt in picked_views]

# Step 2: select reference sheets
# ======================================================================
selected_sheets = collect_selected_sheets(doc, uidoc)
if not selected_sheets:
    all_sheets = collect_all_sheets(doc)
    if not all_sheets:
        output.print_md('**No sheets found in this project.**')
        script.exit()

    sheet_options = [ListOption(s, get_sheet_display_name(s)) for s in all_sheets]
    picked_sheets = forms.SelectFromList.show(
        sheet_options,
        name_attr='display_name',
        multiselect=True,
        title='Select Reference Sheets'
    )
    if not picked_sheets:
        script.exit()
    selected_sheets = [opt.item for opt in picked_sheets]

# Step 3: view template (optional)
# ======================================================================
target_view_template_id = prompt_for_view_template(doc)

# Step 4: sheet number prefix
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

# Step 5: sheet collection (Revit 2024+ only; silently skipped on older versions)
# ======================================================================
collection_element, collection_new_name = prompt_for_sheet_collection(doc)

reference_index = build_reference_viewport_index(doc, selected_sheets)
if not reference_index:
    output.print_md('**No scoped views found on selected reference sheets.**')
    script.exit()

# Existing sheet numbers to avoid collisions
existing_sheet_numbers = {
    s.SheetNumber for s in FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
    if getattr(s, 'SheetNumber', None)
}

created = []
failed = []
used_reference_viewports = set()


def set_view_template_if_requested(target_view, template_id):
    if template_id is not None and template_id != ElementId.InvalidElementId:
        try:
            target_view.ViewTemplateId = template_id
        except Exception:
            pass


with revit.Transaction('Create Sheets'):
    for target_view in selected_views:
        sheet = None
        try:
            scope_key = get_scope_box_key(target_view)
            if not scope_key:
                raise Exception('target view has no scope box')
            scope_suffix = get_scope_box_suffix(doc, target_view)
            if not scope_suffix:
                raise Exception('could not resolve scope box name suffix')

            candidates = reference_index.get(scope_key, [])
            match = choose_best_reference(
                target_view, candidates, used_reference_viewports)
            if not match:
                raise Exception('no matching reference viewport found for scope box')

            source_sheet, source_vp, _ = match
            used_reference_viewports.add(source_vp.Id.IntegerValue)

            source_titleblock_symbol_id = get_sheet_titleblock_symbol_id(doc, source_sheet)
            sheet = ViewSheet.Create(doc, source_titleblock_symbol_id)

            target_sheet_number = get_unique_scope_sheet_number(
                sheet_prefix, scope_suffix, existing_sheet_numbers)
            if not target_sheet_number:
                target_sheet_number = get_unique_scope_sheet_number(
                    sheet_prefix, source_sheet.SheetNumber, existing_sheet_numbers)
            sheet.SheetNumber = target_sheet_number
            sheet.Name = source_sheet.Name

            copy_writable_parameters(
                source_sheet,
                sheet,
                skip_param_names=['Sheet Number', 'Sheet Name']
            )

            source_tb = get_sheet_titleblock_instance(doc, source_sheet)
            target_tb = get_sheet_titleblock_instance(doc, sheet)
            copy_titleblock_instance_parameters(source_tb, target_tb)

            copy_reference_sheet_content(doc, source_sheet, sheet)

            # Ensure prefix-based numbering always wins after any copy operations.
            sheet.SheetNumber = target_sheet_number

            set_view_template_if_requested(target_view, target_view_template_id)

            # Regenerate so sheet Outline and view geometry are valid
            doc.Regenerate()

            center = get_viewport_center(source_vp)
            if center is None:
                center = get_sheet_center(sheet)

            if not Viewport.CanAddViewToSheet(doc, sheet.Id, target_view.Id):
                raise Exception('target view cannot be added to sheet (already placed?)')

            new_vp = Viewport.Create(doc, sheet.Id, target_view.Id, center)
            apply_viewport_layout(source_vp, new_vp)

            created.append((sheet, target_view))
        except Exception as err:
            try:
                if sheet is not None:
                    doc.Delete(sheet.Id)
            except Exception:
                pass
            failed.append('{} ({})'.format(
                get_view_display_name(target_view), err))

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
    for sheet, target_view in created:
        output.print_md('- {} -> {}'.format(
            get_view_display_name(target_view), output.linkify(sheet.Id)))

if failed:
    output.print_md('---')
    output.print_md('## Could not create {} sheet(s)'.format(len(failed)))
    for message in failed:
        output.print_md('- {}'.format(message))
