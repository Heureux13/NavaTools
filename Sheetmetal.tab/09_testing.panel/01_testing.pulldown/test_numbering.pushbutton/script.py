# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.DB import ElementId
from pyrevit import script, revit
from config.parameters_registry import (
    PYT_NUMBER_FABRICATION,
    PYT_NUMBER_ORDER,
    PYT_SKIP_NUMBER,
    RVT_ITEM_NUMBER,
)
from ducts.revit_duct import RevitDuct
from ducts.revit_numbering import RevitNumbers
import ducts.revit_numbering as numbering_module

# Button info
# ======================================================================
__title__ = 'Test Numbering'
__doc__ = '''
Test Numbering
'''

# Variables
# ======================================================================

output = script.get_output()
target_id = 4983007


def _parameter_report(element, parameter_name):
    parameter = element.LookupParameter(parameter_name)
    if not parameter:
        return "missing"

    value = parameter.AsString()
    if value is None:
        value = parameter.AsValueString()
    if value is None:
        value = ""

    return "{} (read-only={})".format(str(value).strip(), parameter.IsReadOnly)


def _family_name(element):
    parameter = element.LookupParameter("Family")
    if not parameter:
        return ""

    value = parameter.AsString()
    if value is None:
        value = parameter.AsValueString()
    return "" if value is None else str(value).strip()


def _connector_report(numbering, duct):
    results = []
    for connector_index, connector in enumerate(duct.get_connectors()):
        try:
            connected = connector.IsConnected
        except Exception as exception:
            results.append("connector {}: IsConnected error {}".format(
                connector_index, exception))
            continue

        refs = []
        try:
            for reference in connector.AllRefs:
                owner = getattr(reference, "Owner", None)
                owner_id = getattr(getattr(owner, "Id", None), "IntegerValue", None)
                owner_family = _family_name(owner) if owner else ""
                refs.append("{} ({})".format(owner_id, owner_family))
        except Exception as exception:
            refs.append("AllRefs error {}".format(exception))

        results.append("connector {}: connected={}, refs={}".format(
            connector_index,
            connected,
            ", ".join(refs) if refs else "none",
        ))

    return results


try:
    doc = revit.doc
    view = revit.active_view
    element = doc.GetElement(ElementId(target_id))

    output.print_md("## Numbering diagnostic for element {}".format(target_id))
    if element is None:
        output.print_md("Element was not found in the active document.")
        script.exit()

    target = RevitDuct(doc, view, element)
    numbering = RevitNumbers(output_obj=output)

    output.print_md("### Element state")
    output.print_md("- Category: {}".format(target.category))
    output.print_md("- Family: `{}`".format(target.family))
    family_value = target.family or ""
    normalized_family = family_value.strip().lower()
    family_codes = [ord(character) for character in family_value]
    output.print_md("- Raw family: `{}`".format(repr(family_value)))
    output.print_md("- Family length: `{}`".format(len(family_value)))
    output.print_md("- Family character codes: `{}`".format(family_codes))
    output.print_md("- Normalized family: `{}`".format(repr(normalized_family)))
    output.print_md("- Direct family membership: `{}`".format(
        normalized_family in numbering.number_families))
    output.print_md("- Numbering module: `{}`".format(
        getattr(numbering_module, "__file__", "unknown")))
    output.print_md("- Type: `{}`".format(_parameter_report(element, "Type")))
    output.print_md("- Item Number: `{}`".format(
        _parameter_report(element, RVT_ITEM_NUMBER)))
    output.print_md("- NumberFabrication: `{}`".format(
        _parameter_report(element, PYT_NUMBER_FABRICATION)))
    output.print_md("- NumberOrder: `{}`".format(
        _parameter_report(element, PYT_NUMBER_ORDER)))
    output.print_md("- SkipNumber: `{}`".format(
        _parameter_report(element, PYT_SKIP_NUMBER)))
    output.print_md("- is_numberable: `{}`".format(numbering.is_numberable(target)))
    output.print_md("- is_traversable: `{}`".format(numbering.is_traversable(target)))
    output.print_md("- has_skip_value: `{}`".format(numbering.has_skip_value(target)))
    output.print_md("- has_stop_value: `{}`".format(numbering.has_stop_value(target)))

    output.print_md("### Connector state")
    connector_lines = _connector_report(numbering, target)
    if connector_lines:
        for line in connector_lines:
            output.print_md("- {}".format(line))
    else:
        output.print_md("- No connectors returned by RevitDuct.get_connectors().")

    direct_neighbors = numbering._get_connected_fittings(target)
    output.print_md("- Numbering neighbors: {}".format(len(direct_neighbors)))
    for neighbor in direct_neighbors:
        output.print_md("  - {}: {}".format(neighbor.id, neighbor.family))

    output.print_md("### Ordered-run reachability")
    ordered_ducts = numbering.get_order_numbers(scope="view")
    output.print_md("- Ordered starts in active view: {}".format(len(ordered_ducts)))

    found_in_runs = []
    for start_duct in ordered_ducts:
        connectivity_map = numbering.build_connectivity_map(start_duct)
        if target_id in connectivity_map:
            found_in_runs.append(start_duct)

    if found_in_runs:
        for start_duct in found_in_runs:
            output.print_md(
                "- Reachable from ordered start {} (order `{}`; item `{}`)".format(
                    start_duct.id,
                    numbering.get_order_number_text(start_duct),
                    numbering.get_item_number(start_duct),
                )
            )
    else:
        output.print_md(
            "- **Not present in any ordered start connectivity map.**"
        )

except Exception as exception:
    output.print_md("## Diagnostic failed: {}".format(exception))
    import traceback
    output.print_md("```text\n{}\n```".format(traceback.format_exc()))
