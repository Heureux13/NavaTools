# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, ElementId, VisibleInViewFilter
from Autodesk.Revit.UI import TaskDialog
from pyrevit import revit, script
from System.Collections.Generic import List
import re
from constants.print_outputs import print_disclaimer
from config.parameters_registry import (
    PYT_NUMBER_ORDER,
    BBM_UNIT,
    BBM_VAV,
    BBM_DUTY,
    PYT_NUMBER_RUN,
)


# Button info
# ===================================================
__title__ = "Select by Order Number"
__doc__ = """
Selects all ducts by accending order number in current view.
"""

# Variables
# ==================================================
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
view = revit.active_view
output = script.get_output()

# Helpers
# ========================================================================


def has_number_order_value(element):
    param = element.LookupParameter(PYT_NUMBER_ORDER)
    if not param:
        return False

    # HasValue is the safest indicator for non-string parameters in Revit.
    try:
        if not param.HasValue:
            return False
    except Exception:
        pass

    # Treat empty strings as no value.
    try:
        text_value = param.AsString()
        if text_value is not None:
            return text_value.strip() != ""
    except Exception:
        pass

    return True


def get_param_display_value(element, param_name):
    param = element.LookupParameter(param_name)
    if not param:
        return "None"

    try:
        if param.AsString() is not None:
            value = param.AsString().strip()
            return value if value else "None"
    except Exception:
        pass

    try:
        if param.AsValueString() is not None:
            value = param.AsValueString().strip()
            return value if value else "None"
    except Exception:
        pass

    try:
        return str(param.AsInteger())
    except Exception:
        pass

    try:
        return str(param.AsDouble())
    except Exception:
        pass

    return "None"


def get_number_order_sort_key(element):
    raw_value = get_param_display_value(element, PYT_NUMBER_ORDER)
    if raw_value == "None":
        return (1, float("inf"), element.Id.IntegerValue)

    text = str(raw_value).strip()

    try:
        return (0, int(text), element.Id.IntegerValue)
    except Exception:
        pass

    # Handle values like "1.0" or formatted strings.
    try:
        return (0, int(float(text)), element.Id.IntegerValue)
    except Exception:
        pass

    match = re.search(r"-?\d+", text)
    if match:
        try:
            return (0, int(match.group(0)), element.Id.IntegerValue)
        except Exception:
            pass

    return (1, float("inf"), element.Id.IntegerValue)


def get_order_number_display(element):
    sort_key = get_number_order_sort_key(element)
    if sort_key[0] == 0:
        return "{:03}".format(int(sort_key[1]))
    return get_param_display_value(element, PYT_NUMBER_ORDER)


def get_whole_order_number(element):
    raw_value = get_param_display_value(element, PYT_NUMBER_ORDER)
    if raw_value == "None":
        return None

    text = str(raw_value).strip()
    if re.match(r"^-?\d+$", text):
        try:
            return int(text)
        except Exception:
            return None

    return None


def format_missing_ranges(numbers):
    if not numbers:
        return "None"

    ranges = []
    start = numbers[0]
    end = numbers[0]

    for n in numbers[1:]:
        if n == end + 1:
            end = n
        else:
            if start == end:
                ranges.append(str(start))
            else:
                ranges.append("{}-{}".format(start, end))
            start = n
            end = n

    if start == end:
        ranges.append(str(start))
    else:
        ranges.append("{}-{}".format(start, end))

    return ", ".join(ranges)


# Main Code
# ==================================================
try:
    # Collect only fabrication ductwork strictly visible in the active view
    fab_duct = (FilteredElementCollector(doc, view.Id)
                .OfCategory(BuiltInCategory.OST_FabricationDuctwork)
                .WhereElementIsNotElementType()
                .WherePasses(VisibleInViewFilter(doc, view.Id))
                .ToElements())

    # Combines list into one (only fab ductwork available)
    all_duct = list(fab_duct)

    if not all_duct:
        TaskDialog.Show("No Ducts", "No ducts found in current view.")
        script.exit()

    duct_run = [d for d in all_duct if has_number_order_value(d)]
    if not duct_run:
        TaskDialog.Show(
            "No Selection",
            "No ducts with a value in {} were found in current view.".format(PYT_NUMBER_ORDER)
        )
        script.exit()

    duct_run = sorted(duct_run, key=get_number_order_sort_key)

    # Select ducts in Revit
    duct_ids = List[ElementId]()
    for d in duct_run:
        duct_ids.Add(d.Id)
    uidoc.Selection.SetElementIds(duct_ids)

    # Final printout with the requested parameter fields
    if len(duct_run) < 500:
        for i, d in enumerate(duct_run, start=1):
            order_number = get_order_number_display(d)
            item_number = get_param_display_value(d, "Item Number")
            bbm_unit = get_param_display_value(d, BBM_UNIT)
            bbm_vav = get_param_display_value(d, BBM_VAV)
            bbm_duty = get_param_display_value(d, BBM_DUTY)
            number_run = get_param_display_value(d, PYT_NUMBER_RUN)

            output.print_md(
                "### No: {:03} | ID: {} | Order #: {} | Item: {} | Unit: {} | VAV: {} | Duty: {} | Run: {}".format(
                    i,
                    output.linkify(
                        d.Id),
                    order_number,
                    item_number,
                    bbm_unit,
                    bbm_vav,
                    bbm_duty,
                    number_run,
                ))

    element_ids = [d.Id for d in duct_run]
    output.print_md("---")
    output.print_md(
        "# Total Elements: {}, {}".format(
            len(duct_run),
            output.linkify(element_ids)
        )
    )

    whole_orders = sorted(set(
        n for n in (get_whole_order_number(d) for d in duct_run)
        if n is not None
    ))

    if whole_orders:
        min_order = whole_orders[0]
        max_order = whole_orders[-1]
        whole_order_set = set(whole_orders)
        missing_orders = [
            n for n in range(min_order, max_order + 1)
            if n not in whole_order_set
        ]

        output.print_md(
            "# Whole Order Range: {} to {}".format(min_order, max_order)
        )
        output.print_md(
            "# Missing Whole Numbers: {}".format(format_missing_ranges(missing_orders))
        )
    else:
        output.print_md("# Missing Whole Numbers: Unable to evaluate (no whole-number Order values found).")

    # Final print statements
    print_disclaimer(output)

except Exception as e:
    TaskDialog.Show("Error", "Script failed: {}".format(str(e)))
    import traceback
    output.print_md("```\n{}\n```".format(traceback.format_exc()))
