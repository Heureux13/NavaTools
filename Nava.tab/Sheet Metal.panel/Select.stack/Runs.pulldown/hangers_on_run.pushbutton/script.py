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
from revit_output import print_parameter_help
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

print_parameter = [
    '_weight_supporting',
    'mark',
]


def safe_float(val):
    try:
        return float(val)
    except Exception:
        return 0.0

# Main Code
# ==================================================


# Get all ducts
ducts = RevitDuct.all(doc, view)
duct = RevitDuct.from_selection(uidoc, doc, view)

# Filter down to short joints
selected_duct = RevitDuct.from_selection(uidoc, doc, view)
selected_duct = selected_duct[0] if selected_duct else None

# Start of select / print loop
if selected_duct:
    # Selets duct that is connected to the selected duct based on size
    run = RevitDuct.create_duct_run(selected_duct, doc, view)
    RevitElement.select_many(uidoc, run)
    total_length = sum(safe_float(RevitDuct.parse_length_string(
        d.centerline_length)) or 0 for d in run)
    total_weight = sum(safe_float(d.weight) or 0 for d in run)

    # Get all hangers in the view and filter for ones supporting ducts in the run
    hangers = []
    duct_ids = set([d.element.Id for d in run])

    # Collect all fabrication hangers in the view
    all_hangers = FilteredElementCollector(doc, view.Id)\
        .OfCategory(BuiltInCategory.OST_FabricationHangers)\
        .WhereElementIsNotElementType()\
        .ToElements()

    # Check which hangers reference our duct size
    duct_size = run[0].size if run else None  # Get size from first duct
    # Normalize duct size for comparison
    duct_size_no_x = duct_size.replace(
        "x", "") if duct_size else None  # 12"x12" -> 12"12"
    duct_size_no_symbol = duct_size.replace("ø", "").replace(
        "x", "") if duct_size else None  # 12"ø -> 12"

    for hanger in all_hangers:
        # Check "Size of Primary End" parameter
        size_param = hanger.LookupParameter("Size of Primary End")
        if size_param:
            hanger_size = size_param.AsString()
            # Normalize hanger size
            hanger_size_clean = hanger_size.replace(
                "ø", "").replace("x", "") if hanger_size else None

            # Match formats: 12"x12", 12"12", 12"ø, 12"
            if (hanger_size == duct_size or
                hanger_size == duct_size_no_x or
                hanger_size == duct_size_no_symbol or
                    hanger_size_clean == duct_size_no_symbol):
                hangers.append(hanger)

    # Select the hangers
    if hangers:
        weight_per_hanger = round(total_weight / len(hangers), 2)
        hanger_ids = [h.Id for h in hangers]
        RevitElement.select_many(uidoc, hangers)
        output.print_md("---")
        output.print_md("# Found {} hangers on the run: {}".format(
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
                for parameter_name in print_parameter:
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

            # Total count
            duct_element_ids = [d.element.Id for d in run]
            output.print_md("---")
            output.print_md("# Duct Run Information")
            output.print_md(
                "### Duct Qty: {:02} | Length: {:06.2f}ft | Run Weight: {:6.2f}lbs | lbs/ft: {:6.2f} | {}".format(
                    len(duct_element_ids),
                    round(total_length / 12, 3),
                    total_weight,
                    total_weight / (total_length / 12),
                    output.linkify(duct_element_ids)
                )
            )

        # Final print statements
        print_parameter_help(output)
else:
    output.print_md("## Select a duct first")
