# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import revit, script
from Autodesk.Revit.DB import BuiltInCategory, FilteredElementCollector, StorageType, View
from Autodesk.Revit.DB import UnitUtils

try:
    from Autodesk.Revit.DB import UnitTypeId
except Exception:
    UnitTypeId = None

try:
    from Autodesk.Revit.DB import DisplayUnitType
except Exception:
    DisplayUnitType = None
from importlib import import_module

# Button info
# ======================================================================
__title__ = 'Weight Per Foot'
__doc__ = '''
Collects Fabrication Ductwork and Fabrication Pipework,
calculates Weight / Length (lbs/ft),
and writes the result into _UMI_PYT_WeightPerFoot.
'''

# Variables
# ======================================================================

doc = revit.doc
# Prefer pyRevit context, then fallback to __revit__ host.
uidoc = getattr(revit, 'uidoc', None)
if uidoc is None:
    revit_host = globals().get('__revit__')
    uidoc = revit_host.ActiveUIDocument if revit_host else None
output = script.get_output()

# Read from config when available; fallback keeps script usable in test contexts.
try:
    registry = import_module('config.parameters_registry')
    PYT_WEIGHT_PER_FOOT = getattr(
        registry, 'PYT_WEIGHT_PER_FOOT', '_UMI_PYT_WeightPerFoot')
except Exception:
    PYT_WEIGHT_PER_FOOT = '_UMI_PYT_WeightPerFoot'

RVT_LENGTH = 'Length'
RVT_WEIGHT = 'Weight'


def collect_category_elements(document, source_view, category):
    return (FilteredElementCollector(document, source_view.Id)
            .OfCategory(category)
            .WhereElementIsNotElementType()
            .ToElements())


def get_element_id_value(element_id):
    if element_id is None:
        return None
    try:
        return int(element_id.IntegerValue)
    except Exception:
        return int(element_id.Value)


def get_selected_views(document, ui_document):
    selected_ids = ui_document.Selection.GetElementIds()
    if not selected_ids:
        return []

    selected_views = []
    for element_id in selected_ids:
        element = document.GetElement(element_id)
        if isinstance(element, View) and not element.IsTemplate:
            selected_views.append(element)
    return selected_views


def lookup_parameter_case_insensitive(element, parameter_name):
    if not element or not parameter_name:
        return None

    target = parameter_name.strip().lower()
    for parameter in element.Parameters:
        try:
            if parameter.Definition and parameter.Definition.Name.strip().lower() == target:
                return parameter
        except Exception:
            continue
    return None


def read_double_value(parameter):
    if not parameter:
        return None

    try:
        if parameter.StorageType == StorageType.Double:
            return parameter.AsDouble()
        if parameter.StorageType == StorageType.Integer:
            return float(parameter.AsInteger())
        if parameter.StorageType == StorageType.String:
            raw = parameter.AsString()
            if raw is None:
                raw = parameter.AsValueString()
            return float(raw) if raw else None
    except Exception:
        return None
    return None


def set_weight_per_foot(element, value):
    destination = element.LookupParameter(PYT_WEIGHT_PER_FOOT)
    if not destination or destination.IsReadOnly:
        return False, "missing/read-only destination parameter"

    try:
        if destination.StorageType == StorageType.Double:
            destination.Set(round(float(value), 1))
            return True, None
        if destination.StorageType == StorageType.String:
            destination.Set("{:.1f}".format(float(value)))
            return True, None
        if destination.StorageType == StorageType.Integer:
            destination.Set(int(round(float(value))))
            return True, None
    except Exception as ex:
        return False, str(ex)

    return False, "unsupported destination storage type"


def mass_internal_to_lbs(internal_mass_value):
    if internal_mass_value is None:
        return None

    try:
        if UnitTypeId is not None:
            return UnitUtils.ConvertFromInternalUnits(internal_mass_value, UnitTypeId.PoundsMass)
    except Exception:
        pass

    try:
        if DisplayUnitType is not None:
            return UnitUtils.ConvertFromInternalUnits(internal_mass_value, DisplayUnitType.DUT_POUNDS_MASS)
    except Exception:
        pass

    # Fallback if API conversion is unavailable: Revit internal mass is kg.
    return float(internal_mass_value) * 2.20462262185


if uidoc is None:
    output.print_md("## Could not access active Revit UI document")
    script.exit()

selected_views = get_selected_views(doc, uidoc)

if not selected_views:
    output.print_md("## Select one or more non-template views in Project Browser, then run again")
    script.exit()

duct_map = {}
pipe_map = {}

for selected_view in selected_views:
    try:
        for duct in collect_category_elements(doc, selected_view, BuiltInCategory.OST_FabricationDuctwork):
            duct_map[get_element_id_value(duct.Id)] = duct

        for pipe in collect_category_elements(doc, selected_view, BuiltInCategory.OST_FabricationPipework):
            pipe_map[get_element_id_value(pipe.Id)] = pipe
    except Exception:
        continue

fab_ducts = list(duct_map.values())
fab_pipes = list(pipe_map.values())
elements = fab_ducts + fab_pipes

if not elements:
    output.print_md("## No fabrication duct/pipe elements found in selected views")
    script.exit()

updated_count = 0
skipped_count = 0
error_count = 0

with revit.Transaction("Set Weight Per Foot"):
    for element in elements:
        weight_param = lookup_parameter_case_insensitive(element, RVT_WEIGHT)
        length_param = lookup_parameter_case_insensitive(element, RVT_LENGTH)

        weight_value = read_double_value(weight_param)
        length_ft = read_double_value(length_param)

        if weight_value is None or length_ft is None or length_ft <= 0:
            skipped_count += 1
            continue

        weight_lbs = mass_internal_to_lbs(weight_value)
        if weight_lbs is None:
            skipped_count += 1
            continue

        weight_per_ft = round(weight_lbs / length_ft, 1)
        did_set, _ = set_weight_per_foot(element, weight_per_ft)

        if did_set:
            updated_count += 1
        else:
            error_count += 1

output.print_md("# Weight Per Foot Update")
output.print_md("### Selected views: {}".format(len(selected_views)))
output.print_md("### Fabrication Ducts in selected views: {}".format(len(fab_ducts)))
output.print_md("### Fabrication Pipes in selected views: {}".format(len(fab_pipes)))
output.print_md("### Updated: {}".format(updated_count))
output.print_md("### Skipped (missing weight/length or zero length): {}".format(skipped_count))
output.print_md("### Failed to write destination parameter: {}".format(error_count))
