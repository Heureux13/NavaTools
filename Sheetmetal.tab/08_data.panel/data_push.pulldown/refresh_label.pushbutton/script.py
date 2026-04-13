# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import revit, script
from Autodesk.Revit.DB import BuiltInCategory, FilteredElementCollector, Transaction
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
__title__ = 'Refresh Label Data'
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


def _get_param_case_insensitive(element, param_name):
    target = (param_name or '').strip().lower()
    if not target or element is None:
        return None
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
        return val.strip() if val else ''
    except Exception:
        return ''


def _resolve_hierarchy_value(element, doc, hierarchy):
    """Return last non-empty value found while iterating hierarchy in order."""
    elem_type = None
    try:
        elem_type = doc.GetElement(element.GetTypeId())
    except Exception:
        elem_type = None

    result = ''
    for param_name in hierarchy:
        value = _get_param_text(_get_param_case_insensitive(element, param_name))
        if not value and elem_type:
            value = _get_param_text(_get_param_case_insensitive(elem_type, param_name))
        if value:
            result = value

    return result


def _resolve_cfm_value(element):
    """Resolve PYT_CFM from BBM CFMs; SA wins when both have values."""
    sa_value = _get_param_text(_get_param_case_insensitive(element, BBM_CFM_SA))
    ea_value = _get_param_text(_get_param_case_insensitive(element, BBM_CFM_EA))

    if sa_value:
        return sa_value
    if ea_value:
        return ea_value
    return ''


def _collect_elements_by_category(doc, categories):
    result = []
    seen = set()

    for bic in categories:
        elements = (
            FilteredElementCollector(doc)
            .OfCategory(bic)
            .WhereElementIsNotElementType()
            .ToElements()
        )
        for elem in elements:
            if elem is None:
                continue
            try:
                elem_id = elem.Id.IntegerValue
            except Exception:
                continue
            if elem_id in seen:
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

t = Transaction(doc, 'Refresh BBM Label From Hierarchy')
t.Start()
try:
    for elem in elements:
        try:
            target_param = _get_param_case_insensitive(elem, BBM_LABEL)
            if not target_param or target_param.IsReadOnly:
                skipped += 1
                continue

            new_value = _resolve_hierarchy_value(elem, doc, HIERARCHY)
            old_value = _get_param_text(target_param)

            if old_value == new_value:
                unchanged += 1
            else:
                target_param.Set(new_value)
                updated += 1

            cfm_target_param = _get_param_case_insensitive(elem, PYT_CFM)
            if not cfm_target_param or cfm_target_param.IsReadOnly:
                cfm_skipped += 1
                continue

            new_cfm_value = _resolve_cfm_value(elem)
            old_cfm_value = _get_param_text(cfm_target_param)

            if old_cfm_value == new_cfm_value:
                cfm_unchanged += 1
                continue

            cfm_target_param.Set(new_cfm_value)
            cfm_updated += 1
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
output.print_md('PYT CFM Skipped (missing/read-only PYT CFM): {}'.format(cfm_skipped))

if errors:
    output.print_md('Errors: {}'.format(len(errors)))
    for elem, reason in errors:
        output.print_md('- ID {}: {}'.format(output.linkify(elem.Id), reason))
