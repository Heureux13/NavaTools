# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import revit, script
from config.parameters_registry import *
from Autodesk.Revit.DB import (
    BuiltInCategory,
    ElementId,
    FilteredElementCollector,
    StorageType,
    Transaction,
)
from System.Collections.Generic import List

# Button info
# ======================================================================
__title__ = 'Copy Parameter Value'
__doc__ = '''
Copy legacy _offset values into PYT offset parameters.
'''

# Configuration
# ======================================================================
parameter_map = {
    PYT_OFFSET_CENTER_H: '_offset_center_h',
    PYT_OFFSET_CENTER_V: '_offset_center_v',
    PYT_OFFSET_TOP: '_offset_top',
    PYT_OFFSET_BOTTOM: '_offset_bottom',
    PYT_OFFSET_RIGHT: '_offset_right',
    PYT_OFFSET_LEFT: '_offset_left',
    PYT_OFFSET_VALUE: '_offset',
}

# Code
# ======================================================================
doc = revit.doc
view = revit.active_view
uidoc = revit.uidoc
output = script.get_output()


def _element_id_value(eid):
    if eid is None:
        return None
    return eid.IntegerValue if hasattr(eid, "IntegerValue") else eid.Value


def _lookup_parameter_case_insensitive(element, param_name):
    try:
        direct_param = element.LookupParameter(param_name)
        if direct_param:
            return direct_param
    except Exception:
        pass

    param_name_lower = param_name.strip().lower()

    for param in element.Parameters:
        try:
            if param.Definition.Name.strip().lower() == param_name_lower:
                return param
        except Exception:
            pass

    try:
        element_type = doc.GetElement(element.GetTypeId())
    except Exception:
        element_type = None

    if element_type:
        for param in element_type.Parameters:
            try:
                if param.Definition.Name.strip().lower() == param_name_lower:
                    return param
            except Exception:
                pass

    return None


def _get_parameter_text(param):
    if param is None:
        return ""

    try:
        if param.StorageType == 0:
            return ""
    except Exception:
        pass

    try:
        value = param.AsString()
        if value:
            return value.strip()
    except Exception:
        pass

    try:
        value = param.AsValueString()
        if value:
            return value.strip()
    except Exception:
        pass

    try:
        if param.StorageType == 1:
            return str(param.AsDouble()).strip()
        if param.StorageType == 2:
            return str(param.AsInteger()).strip()
        if param.StorageType == 3:
            return str(_element_id_value(param.AsElementId())).strip()
    except Exception:
        pass

    return ""


def _set_parameter_from_source(target_param, source_param):
    if not target_param or target_param.IsReadOnly:
        return False

    source_text = _get_parameter_text(source_param)
    if source_text == "":
        return False

    try:
        if target_param.StorageType == StorageType.String:
            return bool(target_param.Set(source_text))

        if target_param.StorageType == StorageType.Double:
            if source_param.StorageType == StorageType.Double:
                return bool(target_param.Set(source_param.AsDouble()))
            return bool(target_param.Set(float(source_text)))

        if target_param.StorageType == StorageType.Integer:
            if source_param.StorageType == StorageType.Integer:
                return bool(target_param.Set(source_param.AsInteger()))
            return bool(target_param.Set(int(float(source_text))))

        if target_param.StorageType == StorageType.ElementId:
            if source_param.StorageType == StorageType.ElementId:
                return bool(target_param.Set(source_param.AsElementId()))
            return False
    except Exception:
        return False

    return False


# Collect all MEP elements
categories = [
    BuiltInCategory.OST_FabricationDuctwork,
]

all_elements = []
for cat in categories:
    all_elements.extend(
        FilteredElementCollector(doc, view.Id)
        .OfCategory(cat)
        .WhereElementIsNotElementType()
        .ToElements()
    )

if not all_elements:
    output.print_md("## No MEP elements found in this view.")
    script.exit()

output.print_md("## Copy Legacy Offsets To PYT Offsets")
output.print_md("Total elements: {}".format(len(all_elements)))

updated_elements = []
missing_source = 0
missing_target = 0
skipped_blank = 0
failed_sets = 0
write_count = 0

t = Transaction(doc, "Copy legacy offset values to PYT")
t.Start()
try:
    for elem in all_elements:
        changed_on_element = False

        for target_name, source_name in parameter_map.items():
            source_param = _lookup_parameter_case_insensitive(elem, source_name)
            if not source_param:
                missing_source += 1
                continue

            target_param = _lookup_parameter_case_insensitive(elem, target_name)
            if not target_param:
                missing_target += 1
                continue

            if _get_parameter_text(source_param) == "":
                skipped_blank += 1
                continue

            if _set_parameter_from_source(target_param, source_param):
                changed_on_element = True
                write_count += 1
            else:
                failed_sets += 1

        if changed_on_element:
            updated_elements.append(elem)

    t.Commit()
except Exception:
    t.RollBack()
    raise

if updated_elements:
    id_list = List[ElementId]([e.Id for e in updated_elements])
    uidoc.Selection.SetElementIds(id_list)

output.print_md("---")
output.print_md("### Updated Elements: {}".format(len(updated_elements)))
output.print_md("### Parameter Writes: {}".format(write_count))
output.print_md("### Missing Source Params: {}".format(missing_source))
output.print_md("### Missing Target Params: {}".format(missing_target))
output.print_md("### Blank Source Values Skipped: {}".format(skipped_blank))
output.print_md("### Failed Writes: {}".format(failed_sets))

if updated_elements:
    output.print_md(
        "### Updated Element IDs: {}".format(
            output.linkify([e.Id for e in updated_elements])
        )
    )
    output.print_md("# Total Elements Selected: {}".format(len(updated_elements)))
else:
    output.print_md("# No elements were updated.")
