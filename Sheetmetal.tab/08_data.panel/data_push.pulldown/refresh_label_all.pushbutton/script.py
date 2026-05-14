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
    BBM_CFM_EA,
    BBM_CFM_SA,
    BBM_LABEL,
    PYT_CFM,
    PYT_LABEL,
    RVT_MARK,
    RVT_TYPE_MARK,
)

# Button info
# ======================================================================
__title__ = 'Refresh Label Data All'
__doc__ = '''
Refresh _UMI_BBM_Label from hierarchy:
Type Mark -> Mark -> _UMI_PYT_Label

Last non-empty value in the hierarchy wins.
Applies to air terminals, mechanical equipment, MEP duct,
fabrication ductwork (including stiffeners), and fabrication hangers.
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


def _get_param_case_insensitive(element, param_name):
    target = (param_name or '').strip().lower()
    if not target or element is None:
        return None

    try:
        direct = element.LookupParameter(param_name)
        if direct:
            return direct
    except Exception:
        pass

    for param in element.Parameters:
        try:
            definition = param.Definition
            name = definition.Name if definition else None
            if name and name.strip().lower() == target:
                return param
        except Exception:
            pass
    return None


def _get_param_text(param):
    if not param:
        return ''
    try:
        val = param.AsString()
        if not val:
            val = param.AsValueString()
        if not val:
            return ''
        cleaned = val.strip()
        if cleaned.lower() in INVALID_TEXT_VALUES:
            return ''
        return cleaned
    except Exception:
        return ''


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


def _resolve_cfm_value(element):
    """Resolve PYT_CFM from BBM CFMs; SA wins when both have values."""
    sa_value = _get_param_text(
        _get_param_case_insensitive(element, BBM_CFM_SA))
    ea_value = _get_param_text(
        _get_param_case_insensitive(element, BBM_CFM_EA))

    if sa_value:
        return sa_value
    if ea_value:
        return ea_value
    return ''


def _try_parse_float(value_text):
    text = (value_text or '').strip()
    if not text:
        return None

    text = text.replace(',', '')
    try:
        return float(text)
    except Exception:
        return None


def _set_param_from_text(param, value_text):
    if not param or param.IsReadOnly:
        return False

    storage_type = param.StorageType

    if storage_type == StorageType.String:
        param.Set(value_text or '')
        return True

    numeric_value = _try_parse_float(value_text)
    if storage_type == StorageType.Double:
        if numeric_value is None:
            return False
        param.Set(float(numeric_value))
        return True

    if storage_type == StorageType.Integer:
        if numeric_value is None:
            return False
        param.Set(int(round(numeric_value)))
        return True

    return False


def _get_element_id_value(element_id):
    """Return ElementId value compatible with both Revit 2025 and 2026+."""
    try:
        return element_id.Value
    except Exception:
        try:
            return element_id.IntegerValue
        except Exception:
            return None


def _collect_elements_by_category(doc, categories):
    result = []
    seen = set()

    for bic in categories:
        try:
            elements = (
                FilteredElementCollector(doc)
                .OfCategory(bic)
                .WhereElementIsNotElementType()
                .ToElements()
            )
        except Exception:
            continue

        for elem in elements:
            if elem is None:
                continue
            try:
                elem_id = _get_element_id_value(elem.Id)
            except Exception:
                continue
            if elem_id is None or elem_id in seen:
                continue
            seen.add(elem_id)
            result.append(elem)

    return result


doc = revit.doc
elements = _collect_elements_by_category(doc, TARGET_CATEGORIES)

if not elements:
    output.print_md('No supported elements found in model.')
    script.exit()

updated = 0
unchanged = 0
skipped = 0
cfm_updated = 0
cfm_unchanged = 0
cfm_skipped = 0
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

                cfm_target_param = _get_param_case_insensitive(elem, PYT_CFM)
                if not cfm_target_param or cfm_target_param.IsReadOnly:
                    cfm_skipped += 1
                    continue

                new_cfm_value = _resolve_cfm_value(elem)
                old_cfm_value = _get_param_text(cfm_target_param)

                if old_cfm_value == new_cfm_value:
                    cfm_unchanged += 1
                    continue

                if _set_param_from_text(cfm_target_param, new_cfm_value):
                    cfm_updated += 1
                else:
                    cfm_skipped += 1
            except Exception as ex:
                errors.append((elem, str(ex)))

        t.Commit()
    except Exception:
        t.RollBack()
        raise

output.print_md('Updated: {}'.format(updated))
output.print_md('Unchanged: {}'.format(unchanged))
output.print_md('Skipped (missing/read-only BBM label): {}'.format(skipped))
output.print_md('PYT CFM Updated: {}'.format(cfm_updated))
output.print_md('PYT CFM Unchanged: {}'.format(cfm_unchanged))
output.print_md(
    'PYT CFM Skipped (missing/read-only PYT CFM): {}'.format(cfm_skipped))

if errors:
    output.print_md('Errors: {}'.format(len(errors)))
    for elem, reason in errors:
        output.print_md('- ID {}: {}'.format(output.linkify(elem.Id), reason))
