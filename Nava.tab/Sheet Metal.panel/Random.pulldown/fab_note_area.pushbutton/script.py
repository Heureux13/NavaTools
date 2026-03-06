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
import re
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


def get_fab_note_value(element):
    parameter = get_fab_note_param(element)
    if not parameter:
        return ""
    return (parameter.AsString() or parameter.AsValueString() or "").strip()


def has_parenthetical_marker(note_value):
    return bool(re.search(r"\([^()]+\)", note_value or ""))


def normalize_marker(marker_value):
    cleaned = (marker_value or "").strip()
    if not cleaned:
        return ""
    if cleaned.startswith("(") and cleaned.endswith(")") and len(cleaned) > 2:
        return cleaned
    return "({})".format(cleaned.strip("() "))


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

    pending_by_service = {}
    for element in ductwork:
        note_value = get_fab_note_value(element)
        if has_parenthetical_marker(note_value):
            continue
        service_name = get_service_name(element)
        if service_name not in pending_by_service:
            pending_by_service[service_name] = []
        pending_by_service[service_name].append(element)

    if not pending_by_service:
        forms.alert(
            "No visible fabrication ductwork is missing a parenthesis marker like (A1).",
            exitscript=True
        )

    selected_services = forms.SelectFromList.show(
        sorted(pending_by_service.keys()),
        title='Select Service(s) Missing (XX) Marker',
        multiselect=True,
        button_name='Next'
    )

    if not selected_services:
        sys.exit(0)

    marker_value = forms.ask_for_string(
        prompt='Enter marker to append (example: A1 or (A1)):',
        default='',
        title='Append Fab Note Marker'
    )

    if marker_value is None:
        sys.exit(0)

    marker_value = normalize_marker(marker_value)
    if not marker_value:
        forms.alert("Marker value is required (example: A1).", exitscript=True)

    updated_count = 0
    read_only_count = 0
    missing_param_count = 0

    t = Transaction(doc, "Append Fabrication Note Marker by Service")
    t.Start()
    try:
        for service_name in selected_services:
            for element in pending_by_service.get(service_name, []):
                parameter = get_fab_note_param(element)
                if not parameter:
                    missing_param_count += 1
                    continue
                if parameter.IsReadOnly:
                    read_only_count += 1
                    continue

                current_note = get_fab_note_value(element)
                if has_parenthetical_marker(current_note):
                    continue

                if current_note:
                    new_note = "{} {}".format(current_note, marker_value)
                else:
                    new_note = marker_value

                if parameter.Set(new_note):
                    updated_count += 1
        t.Commit()
    except Exception:
        t.RollBack()
        raise

    TaskDialog.Show(
        "Fabrication Note Marker Appended",
        "Updated: {}\nRead-only skipped: {}\nMissing parameter: {}".format(
            updated_count,
            read_only_count,
            missing_param_count
        )
    )

except Exception as e:
    forms.alert("Error: {}".format(e), exitscript=True)
