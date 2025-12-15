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

hanger_parameters = {
    '_umi_duct_supporting_weight',
    'mark',
}

duct_parameters = {
    '_umi_duct_run_weight',
    'mark'
}

# Main Code
# =================================================
# Get all ducts
duct = RevitDuct.from_selection(uidoc, doc, view)

# Filter down to short joints
selected_duct = RevitDuct.from_selection(uidoc, doc, view)
selected_duct = selected_duct[0] if selected_duct else None

# Start of select / print loop
if selected_duct:
    # Selets duct that is connected to the selected duct based on size
    run = RevitDuct.create_duct_run(selected_duct, doc, view)
    RevitElement.select_many(uidoc, run)
    run_total_length = sum(d.length or 0 for d in run)
    run_total_weight = sum(d.weight or 0 for d in run)

    # Get all hangers that intersect with any duct in the run
    hangers = set()  # Use set to avoid duplicates

    for duct in run:
        bbox = duct.element.get_BoundingBox(None)
        if bbox:
            # Create filter for elements intersecting this bounding box
            outline = Outline(bbox.Min, bbox.Max)
            bbox_filter = BoundingBoxIntersectsFilter(outline)

            # Collect hangers intersecting this duct
            intersecting_hangers = FilteredElementCollector(doc)\
                .OfCategory(BuiltInCategory.OST_FabricationHangers)\
                .WherePasses(bbox_filter)\
                .WhereElementIsNotElementType()\
                .ToElements()

            for h in intersecting_hangers:
                hangers.add(h)

    hangers = list(hangers)  # Convert back to list

    # Select the hangers
    if hangers:
        weight_per_hanger = run_total_weight / len(hangers)
        hanger_ids = [h.Id for h in hangers]
        RevitElement.select_many(uidoc, hangers)
        output.print_md("---")
        output.print_md("### Found {} hangers on the run: {}".format(
            len(hangers),
            output.linkify(hanger_ids)
        ))

        # Print hanger info and set Mark parameter in a transaction
        with revit.Transaction("Set Hanger Mark"):

            for i, h in enumerate(hangers, start=1):
                family_name = doc.GetElement(
                    h.GetTypeId()).FamilyName if hasattr(
                    doc.GetElement(
                        h.GetTypeId()),
                    'FamilyName') else "Unknown"
                output.print_md(
                    "### {} | ID: {} | Supporting: {:6.2f}lbs".format(
                        i,
                        output.linkify(h.Id),
                        weight_per_hanger,
                    ))

                # Set the _weight_supporting parameter on the instance
                set_parameter = None
                for parameter_name in hanger_parameters:
                    p = h.LookupParameter(parameter_name)
                    if not p:
                        # output.print_md(
                        #     'Parameter "{}" not found on hanger ID {}'.format(
                        #         parameter_name,
                        #         output.linkify(h.Id)
                        #     )
                        # )
                        continue
                    elif p.IsReadOnly:
                        # output.print_md(
                        #     'Parameter "{}" is read-only on hanger ID {}'.format(
                        #         parameter_name,
                        #         output.linkify(h.Id)
                        #     )
                        # )
                        continue
                    else:
                        set_parameter = p
                        break

                if set_parameter:
                    set_parameter.Set(weight_per_hanger)
                    # output.print_md(
                    #     'Set parameter "{}" on hanger ID {} to {:6.2f}'.format(
                    #         set_parameter.Definition.Name,
                    #         output.linkify(h.Id),
                    #         weight_per_hanger
                    #     )
                    # )
                else:
                    output.print_md(
                        'Could not set parameter on hanger ID {}'.format(
                            output.linkify(h.Id)
                        )
                    )

            # Set run weight on each duct in the run
            for d in run:
                set_parameter = None
                for parameter_name in duct_parameters:
                    p = d.element.LookupParameter(parameter_name)
                    if not p:
                        continue
                    elif p.IsReadOnly:
                        continue
                    else:
                        set_parameter = p
                        break

                if set_parameter:
                    set_parameter.Set(run_total_weight)

            # Total count
            duct_element_ids = [d.element.Id for d in run]
            total_length_ft = run_total_length / 12.0 if run_total_length else 0.0
            lbs_per_ft = (
                run_total_weight / total_length_ft) if total_length_ft else 0.0
            output.print_md("---")
            output.print_md("# Duct Run Information")
            output.print_md(
                "### Duct Qty: {} | Length: {:06.2f}ft | Run Weight: {:6.2f}lbs | lbs/ft: {:6.2f} | {}".format(
                    len(duct_element_ids),
                    total_length_ft,
                    run_total_weight,
                    lbs_per_ft,
                    output.linkify(duct_element_ids)
                )
            )

        # Final print statements
        print_disclaimer(output)
else:
    output.print_md("## Select a duct first")
