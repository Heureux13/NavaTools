# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import revit, script
from Autodesk.Revit.DB import BuiltInCategory, FilteredElementCollector, Transaction

RVT_TYPE_MARK = 'Type Mark'
RVT_MARK = 'Mark'
BBM_LABEL = '_UMI_BBM_Label'
PYT_LABEL = "_UMI_PYT_Label"

# Button info
# ======================================================================
__title__ = 'Update Label'
__doc__ = '''
Update BBM Label from a hierarchy (highest priority first):
Type Mark  ->  Mark  ->  PYT Label
'''

# Variables
# ======================================================================

output = script.get_output()


def _get_param_case_insensitive(element, param_name):
    target = (param_name or '').strip().lower()
    if not target:
        return None
    for param in element.Parameters:
        try:
            definition = param.Definition
            if definition and definition.Name and definition.Name.strip().lower() == target:
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


def _get_hierarchy_value(element, doc, hierarchy):
    """Check each parameter in order. Each one overrides the previous if it has a value.
    Returns the last non-empty value found, or '' if all are empty."""
    elem_type = None
    try:
        elem_type = doc.GetElement(element.GetTypeId())
    except Exception:
        elem_type = None

    result = ''
    for param_name in hierarchy:
        val = ''
        param = _get_param_case_insensitive(element, param_name)
        val = _get_param_text(param)

        if not val and elem_type:
            param = _get_param_case_insensitive(elem_type, param_name)
            val = _get_param_text(param)

        if val:
            result = val

    return result


def _collect_supported_elements(doc):
    categories = [
        BuiltInCategory.OST_DuctTerminal,
        BuiltInCategory.OST_FabricationDuctwork,
        BuiltInCategory.OST_MechanicalEquipment,
    ]

    seen = set()
    result = []

    for bic in categories:
        elems = (
            FilteredElementCollector(doc)
            .OfCategory(bic)
            .WhereElementIsNotElementType()
            .ToElements()
        )
        for elem in elems:
            if not elem:
                continue
            elem_id = elem.Id.IntegerValue
            if elem_id in seen:
                continue
            seen.add(elem_id)
            result.append(elem)

    return result


doc = revit.doc

elements = _collect_supported_elements(doc)
if not elements:
    output.print_md('No supported elements found in model.')
    script.exit()

hierarchy = [RVT_TYPE_MARK, RVT_MARK, PYT_LABEL]

updated = 0
skipped = 0
errors = []

t = Transaction(doc, 'Update BBM Label From Hierarchy')
t.Start()
try:
    for elem in elements:
        try:
            target_param = _get_param_case_insensitive(elem, BBM_LABEL)
            if not target_param:
                skipped += 1
                continue

            if target_param.IsReadOnly:
                skipped += 1
                continue

            value = _get_hierarchy_value(elem, doc, hierarchy)
            target_param.Set(value)
            updated += 1
        except Exception as ex:
            errors.append((elem, str(ex)))

    t.Commit()
except Exception:
    t.RollBack()
    raise

output.print_md('Updated: {}'.format(updated))
output.print_md('Skipped: {}'.format(skipped))

if errors:
    output.print_md('Errors: {}'.format(len(errors)))
    for elem, reason in errors:
        output.print_md('- ID {}: {}'.format(output.linkify(elem.Id), reason))
