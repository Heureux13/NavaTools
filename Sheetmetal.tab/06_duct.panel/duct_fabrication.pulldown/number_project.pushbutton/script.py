# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import revit, script

from config.parameters_registry import PYT_NUMBER_ORDER
from ducts.revit_numbering import RevitNumbers

# Button info
# ======================================================================
__title__ = 'Number Project'
__doc__ = '''
Numbers order duct in entire project
'''

# Variables
# ======================================================================

output = script.get_output()


def _order_value(duct):
    param = duct.element.LookupParameter(PYT_NUMBER_ORDER)
    if not param:
        return ""

    value = param.AsString()
    if value is None:
        value = param.AsValueString()

    return "" if value is None else str(value).strip()


try:
    output.print_md("## Starting numbering...")

    try:
        output.print_md("- Initializing RevitNumbers")
        numbering = RevitNumbers(output_obj=output)
        output.print_md("- RevitNumbers initialized")
    except Exception as ex:
        output.print_md("## Init failed: {}".format(ex))
        script.exit()

    try:
        output.print_md(
            "- Collecting ducts with order numbers (entire project)")
        ordered_ducts = numbering.get_order_numbers(scope="project")
        output.print_md(
            "- Found {} ducts".format(len(ordered_ducts) if ordered_ducts else 0))
    except Exception as ex:
        output.print_md("## Collection failed: {}".format(ex))
        script.exit()

    if not ordered_ducts:
        output.print_md("## no ducts with numbers found")
        script.exit()

    duplicate_orders = numbering.find_duplicate_order_numbers(ordered_ducts)
    if duplicate_orders:
        output.print_md(
            "## Duplicate order numbers found. Please fix these before running numbering:")
        for order_value in sorted(duplicate_orders.keys()):
            dup_ducts = duplicate_orders[order_value]
            output.print_md("- Order **{}** appears {} times".format(
                order_value,
                len(dup_ducts),
            ))
            for dup_duct in dup_ducts:
                output.print_md(
                    "  - {}".format(output.linkify(dup_duct.element.Id)))
        script.exit()

    if not numbering.has_any_item_number(ordered_ducts):
        output.print_md("## no item number found")
        script.exit()

    try:
        output.print_md("- Starting transaction")

        output.print_md("- Running number_ordered_runs")
        for i, duct in enumerate(ordered_ducts):
            output.print_md(
                "  - Processing duct {} of {}".format(i + 1, len(ordered_ducts)))

        with revit.Transaction("Number Ordered Duct Runs (Project)"):
            results = numbering.number_ordered_runs(
                ordered_ducts,
                repeat_numbers=False,
            )
        output.print_md("- Numbering complete (transaction)")
    except Exception as ex:
        output.print_md("## Transaction failed: {}".format(ex))
        import traceback
        output.print_md("- Traceback: {}".format(traceback.format_exc()))
        script.exit()

    if not results:
        output.print_md("## No valid ordered runs were numbered.")
        script.exit()

    output.print_md("## ✓ Success")
    output.print_md("- Runs numbered: {}".format(len(results)))

    for index, (start_duct, start_number, end_number) in enumerate(results, start=1):
        output.print_md(
            "- Run {} | Order: {} | Start duct: {} | Item start: {} | Item end: {}".format(
                index,
                _order_value(start_duct),
                output.linkify(start_duct.element.Id),
                start_number,
                end_number,
            )
        )

except Exception as ex:
    output.print_md("## Numbering failed")
    output.print_md("- Error: {}".format(ex))
    import traceback
    output.print_md("- Traceback: {}".format(traceback.format_exc()))
