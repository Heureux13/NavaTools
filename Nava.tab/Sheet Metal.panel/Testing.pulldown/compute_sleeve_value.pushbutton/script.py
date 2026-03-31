# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import revit, script
from Autodesk.Revit.DB import BuiltInCategory, FabricationPart, FilteredElementCollector, StorageType

# Button info
# ======================================================================
__title__ = 'Compute Sleeve Value'
__doc__ = '''
Numbers fabrication duct sleeves in the active view.
'''

# Variables
# ======================================================================
uidoc = getattr(revit, 'uidoc', None)
doc = getattr(revit, 'doc', None)
view = getattr(revit, 'active_view', None)

if uidoc is None or doc is None or view is None:
    revit_host = globals().get('__revit__')
    if revit_host is not None:
        uidoc = uidoc or revit_host.ActiveUIDocument
        if uidoc is not None:
            doc = doc or uidoc.Document
            view = view or uidoc.ActiveView

if uidoc is None:
    raise RuntimeError('Unable to access the active Revit UI document.')
if doc is None:
    raise RuntimeError('Unable to access the active Revit document.')
if view is None:
    raise RuntimeError('Unable to access the active Revit view.')

output = script.get_output()

TYPE_PARAM = '_type'
NUMBER_PARAM = '_#'
SLEEVE_VALUE = 'sleeve'


def _element_id_value(element_or_id):
    try:
        return element_or_id.Id.Value
    except AttributeError:
        pass

    try:
        return element_or_id.Id.IntegerValue
    except AttributeError:
        pass

    try:
        return element_or_id.Value
    except AttributeError:
        return element_or_id.IntegerValue


def _get_param_string(element, param_name):
    param = element.LookupParameter(param_name)
    if not param:
        return None

    try:
        value = param.AsString()
        if value is None:
            value = param.AsValueString()
        return value
    except Exception:
        return None


def _normalized_param_value(element, param_name):
    value = _get_param_string(element, param_name)
    if value is None:
        return ''
    return value.strip().lower()


def _set_param_value(element, param_name, value):
    param = element.LookupParameter(param_name)
    if not param or param.IsReadOnly:
        return False

    try:
        if param.StorageType == StorageType.String:
            param.Set(str(value))
            return True
        if param.StorageType == StorageType.Integer:
            param.Set(int(value))
            return True
        if param.StorageType == StorageType.Double:
            param.Set(float(value))
            return True
    except Exception:
        return False

    return False


all_ducts = list(
    FilteredElementCollector(doc)
    .OfClass(FabricationPart)
    .WhereElementIsNotElementType()
    .ToElements()
)

sleeves = []
missing_type = []

for element in all_ducts:
    category = getattr(element, 'Category', None)
    if category is None:
        continue
    if _element_id_value(category.Id) != int(BuiltInCategory.OST_FabricationDuctwork):
        continue

    type_value = _normalized_param_value(element, TYPE_PARAM)

    if not type_value:
        missing_type.append(element)
        continue

    if type_value == SLEEVE_VALUE:
        sleeves.append(element)

if not sleeves:
    output.print_md('## No sleeve elements found in the active view.')
    output.print_md('Fabrication ductwork checked: {}'.format(len(all_ducts)))
    if missing_type:
        output.print_md(
            'Elements missing `_type`: {}'.format(len(missing_type)))
    script.exit()

sleeves.sort(key=_element_id_value)

numbered = []
failed = []

with revit.Transaction('Number sleeve elements'):
    for index, element in enumerate(sleeves, 1):
        if _set_param_value(element, NUMBER_PARAM, index):
            numbered.append(element)
        else:
            failed.append(element)

output.print_md('## Sleeve numbering complete.')
output.print_md('Fabrication ductwork checked: {}'.format(len(all_ducts)))
output.print_md('Sleeves found: {}'.format(len(sleeves)))
output.print_md('Numbered in `_#`: {}'.format(len(numbered)))

if missing_type:
    output.print_md('Elements missing `_type`: {}'.format(len(missing_type)))
if failed:
    output.print_md('Failed to write `_#`: {}'.format(len(failed)))
