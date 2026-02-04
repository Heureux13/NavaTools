# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from revit_element import RevitElement
from revit_duct import RevitDuct
from revit_output import print_disclaimer
from pyrevit import revit, script
from Autodesk.Revit.DB import *

# Button info
# ===================================================
__title__ = "Hangers on Run"
__doc__ = """
Total weight of run / hanger amount.
"""

# Variables
# ==================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
view = revit.active_view
output = script.get_output()

# Ordered preference: custom param first, then Revit built-in Mark, then lowercase fallback
hanger_parameters = [
    '_hang_weight_supporting',
    'Mark',
    'mark',
]

duct_parameters = [
    '_duct_weight_run',
    'Mark',
    'mark',
]

# Main Code
# =================================================

# Filter down to a single selected duct
selected_duct = RevitDuct.from_selection(uidoc, doc, view)
selected_duct = selected_duct[0] if selected_duct else None

if selected_duct:
    # Build the run from the selected duct
    run = RevitDuct.create_duct_run(selected_duct, doc, view)
    RevitElement.select_many(uidoc, run)

    run_total_length = sum(d.length or 0 for d in run)
    run_total_weight = sum(d.weight or 0 for d in run)

    # Collect hangers that intersect any duct in the run
    hangers = set()
    for duct in run:
        bbox = duct.element.get_BoundingBox(None)
        if not bbox:
            continue
        outline = Outline(bbox.Min, bbox.Max)
        bbox_filter = BoundingBoxIntersectsFilter(outline)
        intersecting = (FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_FabricationHangers)
                        .WherePasses(bbox_filter)
                        .WhereElementIsNotElementType()
                        .ToElements())
        for h in intersecting:
            hangers.add(h)

    hangers = list(hangers)

    if hangers:
        weight_per_hanger = run_total_weight / \
            float(len(hangers)) if hangers else 0.0
        hanger_ids = [h.Id for h in hangers]
        RevitElement.select_many(uidoc, hangers)

        output.print_md("---")
        output.print_md("### Found {} hangers on the run: {}".format(
            len(hangers), output.linkify(hanger_ids)
        ))

        # Write parameters
        with revit.Transaction("Set Hanger Mark"):
            # Hanger instance values
            for i, h in enumerate(hangers, start=1):
                output.print_md(
                    "### {} | ID: {} | Supporting: {:6.2f}lbs".format(
                        i, output.linkify(h.Id), weight_per_hanger
                    )
                )

                set_parameter = None
                for name in hanger_parameters:
                    p = h.LookupParameter(name)
                    if not p or p.IsReadOnly:
                        continue
                    set_parameter = p
                    break

                if set_parameter:
                    try:
                        if set_parameter.StorageType == StorageType.Double:
                            set_parameter.Set(weight_per_hanger)
                        elif set_parameter.StorageType == StorageType.String:
                            set_parameter.Set(str(round(weight_per_hanger, 2)))
                    except Exception:
                        pass

            # Run weight on each duct in the run
            for d in run:
                set_parameter = None
                for name in duct_parameters:
                    p = d.element.LookupParameter(name)
                    if not p or p.IsReadOnly:
                        continue
                    set_parameter = p
                    break

                if set_parameter:
                    try:
                        if set_parameter.StorageType == StorageType.Double:
                            set_parameter.Set(round(run_total_weight, 2))
                        elif set_parameter.StorageType == StorageType.String:
                            set_parameter.Set(str(round(run_total_weight, 2)))
                    except Exception:
                        pass

        # Summary
        duct_element_ids = [d.element.Id for d in run]
        total_length_ft = run_total_length / 12.0 if run_total_length else 0.0
        lbs_per_ft = (run_total_weight /
                      total_length_ft) if total_length_ft else 0.0
        output.print_md("---")
        output.print_md("# Duct Run Information")
        output.print_md(
            "### Duct Qty: {} | Length: {:06.2f}ft | Run Weight: {:6.2f}lbs | lbs/ft: {:6.2f} | {}".format(
                len(duct_element_ids), total_length_ft, run_total_weight, lbs_per_ft, output.linkify(
                    duct_element_ids)
            )
        )

        # Final print statements
        print_disclaimer(output)
else:
    output.print_md("## Select a duct first")
