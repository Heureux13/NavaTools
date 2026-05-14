# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import revit, script
from Autodesk.Revit.DB import BuiltInCategory, FilteredElementCollector, StorageType, Transaction
from config.parameters_registry import (
    BBM_LABEL,
    PYT_LABEL,
    RVT_MARK,
    RVT_TYPE_MARK,
)

# Button info
# ======================================================================
__title__ = 'Refresh Label Data'
__doc__ = '''
Refresh _UMI_BBM_Label from hierarchy:
Type Mark -> Mark -> _UMI_PYT_Label

Last non-empty value in the hierarchy wins.
'''

# Variables
# ======================================================================

output = script.get_output()


MAX_TRANSACTION_BATCH_SIZE = 1000

TARGET_CATEGORIES = (
    BuiltInCategory.OST_DuctTerminal,
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_DuctCurves,
    BuiltInCategory.OST_FabricationDuctwork,
    BuiltInCategory.OST_FabricationHangers,

)

HIERARCHY = (
    RVT_TYPE_MARK,
    RVT_MARK,
    PYT_LABEL,
)


INVALID_TEXT_VALUES = {
    '',
    '**',
    'none',
}


def _get_element_id_value(element_id):
    if element_id is None:
        return None

    try:
        return element_id.Value
    except Exception:
        pass

    try:
        return element_id.IntegerValue
    except Exception:
        return None


def _get_param_case_insensitive(element, param_name):
    if not param_name or element is None:
        return None

    try:
        return element.LookupParameter(param_name)
    except Exception:
        return None


def _read_param_as_text(param):
    if not param:
        return ''
    try:
        storage_type = param.StorageType
        if storage_type == StorageType.String:
            return (param.AsString() or '').strip()
        if storage_type == StorageType.Integer:
            return str(param.AsInteger())
        if storage_type == StorageType.Double:
            return str(param.AsDouble())
        return ''
    except Exception:
        return ''


def _normalize_text(value):
    cleaned = (value or '').strip()
    if cleaned.lower() in INVALID_TEXT_VALUES:
        return ''
    return cleaned


def _get_param_text(param):
    return _normalize_text(_read_param_as_text(param))


def _resolve_hierarchy_value(element, doc, hierarchy, type_cache):
    """Return last non-empty value found while iterating hierarchy in order."""
    elem_type = None
    try:
        type_id = element.GetTypeId()
        type_key = _get_element_id_value(type_id)
        if type_key is not None:
            if type_key not in type_cache:
                type_cache[type_key] = doc.GetElement(type_id)
            elem_type = type_cache[type_key]
    except Exception:
        elem_type = None

    result = ''
    for param_name in hierarchy:
        value = _get_param_text(
            _get_param_case_insensitive(element, param_name))
        if not value and elem_type:
            value = _get_param_text(
                _get_param_case_insensitive(elem_type, param_name))
        if value:
            result = value

    return result


def _set_param_from_text(param, value_text):
    if not param or param.IsReadOnly:
        return False

    storage_type = param.StorageType

    if storage_type != StorageType.String:
        return False

    try:
        param.Set(value_text or '')
        return True
    except Exception:
        return False


def _collect_active_view_elements(doc):
    result = []
    active_view = revit.active_view
    if not active_view:
        return result

    try:
        collector = (
            FilteredElementCollector(doc, active_view.Id)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        return result

    for elem in collector:
        result.append(elem)

    return result


doc = revit.doc
elements = _collect_active_view_elements(doc)
if not elements:
    output.print_md('No elements found in active view.')
    script.exit()

output.print_md(
    'Processing all elements in active view ({}).'.format(len(elements)))

updated = 0
unchanged = 0
skipped = 0
errors = []
type_cache = {}

for batch_start in range(0, len(elements), MAX_TRANSACTION_BATCH_SIZE):
    batch = elements[batch_start: batch_start + MAX_TRANSACTION_BATCH_SIZE]
    t = Transaction(doc, 'Refresh BBM Label From Hierarchy')
    t.Start()
    try:
        for elem in batch:
            try:
                target_param = _get_param_case_insensitive(elem, BBM_LABEL)
                if not target_param or target_param.IsReadOnly:
                    skipped += 1
                    continue

                new_value = _resolve_hierarchy_value(
                    elem, doc, HIERARCHY, type_cache)
                old_value = _get_param_text(target_param)

                if old_value == new_value:
                    unchanged += 1
                else:
                    if _set_param_from_text(target_param, new_value):
                        updated += 1
                    else:
                        skipped += 1
            except Exception as ex:
                errors.append((elem, str(ex)))

        t.Commit()
    except Exception:
        t.RollBack()
        raise

output.print_md('Updated: {}'.format(updated))
output.print_md('Unchanged: {}'.format(unchanged))
output.print_md('Skipped (missing/read-only BBM label): {}'.format(skipped))

if errors:
    output.print_md('Errors: {}'.format(len(errors)))
    for elem, reason in errors:
        try:
            elem_id = _get_element_id_value(elem.Id)
        except Exception:
            elem_id = 'Unknown'
        output.print_md('- ID {}: {}'.format(elem_id, reason))
