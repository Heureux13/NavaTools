# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    FabricationPart,
    VisibleInViewFilter,
    Transaction,
)
from Autodesk.Revit.UI import TaskDialog
from pyrevit import revit, forms
import sys


# Button info
# ===================================================
__title__ = "Add fab note"
__doc__ = """
dddddddddddd
"""

# Variables
# ==================================================
doc = revit.doc
view = revit.active_view


def get_fab_note_param(element):
    for param_name in ("Fabrication Notes", "Fabrication Note"):
        parameter = element.LookupParameter(param_name)
        if parameter:
            return parameter
    return None


def is_blank_fab_note(element):
    parameter = get_fab_note_param(element)
    if not parameter:
        return False
    value = parameter.AsString() or parameter.AsValueString() or ""
    return value.strip() == ""


def get_service_name(element):
    try:
        service_name = element.ServiceName
        if service_name and service_name.strip():
            return service_name.strip()
    except Exception:
        pass

    for param_name in ("Fabrication Service", "Fabrication Service Name", "Service"):
        parameter = element.LookupParameter(param_name)
        if not parameter:
            continue
        value = parameter.AsString() or parameter.AsValueString() or ""
        if value.strip():
            return value.strip()

    return "(No Service)"


# Main Code
# ==================================================
try:
    ductwork = (FilteredElementCollector(doc, view.Id)
                .OfClass(FabricationPart)
                .OfCategory(BuiltInCategory.OST_FabricationDuctwork)
                .WhereElementIsNotElementType()
                .WherePasses(VisibleInViewFilter(doc, view.Id))
                .ToElements())

    if not ductwork:
        forms.alert("No fabrication ductwork found in the active view.", exitscript=True)

    empty_by_service = {}
    for element in ductwork:
        if not is_blank_fab_note(element):
            continue
        service_name = get_service_name(element)
        if service_name not in empty_by_service:
            empty_by_service[service_name] = []
        empty_by_service[service_name].append(element)

    if not empty_by_service:
        forms.alert(
            "No visible fabrication ductwork has an empty Fabrication Note.",
            exitscript=True
        )

    selected_services = forms.SelectFromList.show(
        sorted(empty_by_service.keys()),
        title='Select Service(s) with Empty Fabrication Note',
        multiselect=True,
        button_name='Next'
    )

    if not selected_services:
        sys.exit(0)

    note_value = forms.ask_for_string(
        prompt='Enter Fabrication Note to apply to selected service(s):',
        default='',
        title='Set Fabrication Note'
    )

    if note_value is None:
        sys.exit(0)

    note_value = note_value.strip()
    if not note_value:
        forms.alert("Fabrication Note value is required.", exitscript=True)

    updated_count = 0
    read_only_count = 0
    missing_param_count = 0

    t = Transaction(doc, "Set Fabrication Note by Service")
    t.Start()
    try:
        for service_name in selected_services:
            for element in empty_by_service.get(service_name, []):
                parameter = get_fab_note_param(element)
                if not parameter:
                    missing_param_count += 1
                    continue
                if parameter.IsReadOnly:
                    read_only_count += 1
                    continue
                if not is_blank_fab_note(element):
                    continue
                if parameter.Set(note_value):
                    updated_count += 1
        t.Commit()
    except Exception:
        t.RollBack()
        raise

    TaskDialog.Show(
        "Fabrication Note Updated",
        "Updated: {}\nRead-only skipped: {}\nMissing parameter: {}".format(
            updated_count,
            read_only_count,
            missing_param_count
        )
    )

except Exception as e:
    pass
