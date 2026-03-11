# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from pyrevit import revit, script, DB
from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    FilteredElementCollector,
    FamilySymbol,
    IndependentTag,
    Transaction,
    XYZ,
)
from revit_tagging import RevitTagging

# Button display information
# =================================================
__title__ = "Check GRDs"
__doc__ = """
Prints out quantity and labels of GRDs in the current view.
"""

# Parameters to check for GRDs
# ==================================================
grd_label_param = [
    '_grd_label',
]

# Possible system abbreviation parameter names (checked in order)
service_param_names = [
    'system abbreviation',
    'system abbr',
]


def _get_param_value_case_insensitive(element, doc, param_name):
    param_name_lower = param_name.lower().strip()

    # Instance parameters
    for param in element.Parameters:
        if param.Definition.Name.lower().strip() == param_name_lower:
            try:
                val = param.AsString()
                if not val:
                    val = param.AsValueString()
                return (val or "").strip()
            except Exception:
                return ""

    # Type parameters
    elem_type = doc.GetElement(element.GetTypeId())
    if elem_type:
        for param in elem_type.Parameters:
            if param.Definition.Name.lower().strip() == param_name_lower:
                try:
                    val = param.AsString()
                    if not val:
                        val = param.AsValueString()
                    return (val or "").strip()
                except Exception:
                    return ""

    return ""


def _get_first_value_from_list(element, doc, param_list):
    for param_name in param_list:
        val = _get_param_value_case_insensitive(element, doc, param_name)
        if val:
            return val
    return ""


# Code
# ==================================================
doc = revit.doc
view = revit.active_view
output = script.get_output()

air_terminals = (
    FilteredElementCollector(doc, view.Id)
    .OfCategory(BuiltInCategory.OST_DuctTerminal)
    .WhereElementIsNotElementType()
    .ToElements()
)

if not air_terminals:
    output.print_md("## No air terminals found in this view.")
    script.exit()

empty_label = "(empty)"

output.print_md("## GRD Counts (Current View)")
for param_name in grd_label_param:
    value_counts = {}
    value_ids = {}
    for elem in air_terminals:
        value = _get_param_value_case_insensitive(elem, doc, param_name)
        key = value if value else empty_label
        value_counts[key] = value_counts.get(key, 0) + 1
        value_ids.setdefault(key, []).append(elem.Id)

    output.print_md("\n### {}".format(param_name))
    for value in sorted(value_counts.keys()):
        links = ", ".join(output.linkify(eid) for eid in value_ids.get(value, []))
        output.print_md("- **{}**: {} | IDs: {}".format(value, value_counts[value], links))

output.print_md("\n**Total Air Terminals:** {}".format(len(air_terminals)))

# Service breakdown: service -> label counts
label_param = grd_label_param[0] if grd_label_param else '_grd_label'
service_label_counts = {}
service_label_ids = {}
for elem in air_terminals:
    label = _get_param_value_case_insensitive(elem, doc, label_param)
    label_key = label if label else empty_label
    service = _get_first_value_from_list(elem, doc, service_param_names)
    service_key = service if service else empty_label

    svc_counts = service_label_counts.setdefault(service_key, {})
    svc_counts[label_key] = svc_counts.get(label_key, 0) + 1

    svc_ids = service_label_ids.setdefault(service_key, {})
    svc_ids.setdefault(label_key, []).append(elem.Id)

output.print_md("\n## GRD Counts by Fabrication Service (Current View)")
for service_value in sorted(service_label_counts.keys()):
    output.print_md("\n### {}".format(service_value))
    for label_value in sorted(service_label_counts[service_value].keys()):
        count = service_label_counts[service_value][label_value]
        links = ", ".join(
            output.linkify(eid)
            for eid in service_label_ids.get(service_value, {}).get(label_value, [])
        )
        output.print_md("- **{}**: {} | IDs: {}".format(label_value, count, links))
