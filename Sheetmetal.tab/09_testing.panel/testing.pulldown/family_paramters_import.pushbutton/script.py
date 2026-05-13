# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from revit.revit_element import RevitElement
from ducts.revit_duct import RevitDuct
from ducts.revit_hanger import RevitHanger
from runs.revit_runs import RevitRuns
from constants.print_outputs import print_disclaimer
from pyrevit import revit, script
from Autodesk.Revit.DB import *
from config.parameters_registry import *

# Button info
# ===================================================
__title__ = "Place Hangers on Run"
__doc__ = """
Place hangers on straight/spiral duct runs.
8 feet spacing, 6" from ends, positioned below duct.
"""

# Variables
# ==================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
view = revit.active_view
output = script.get_output()

# Configuration
ALLOWED_FAMILIES = ["Straight", "Spiral Duct"]
HANGER_SPACING_INCHES = 96.0  # 8 feet
END_OFFSET_INCHES = 6.0
HANGER_BELOW_ELEVATION_INCHES = 6.0  # 6" below lower elevation

# Main Code
# =================================================

# Get selected duct
selected_duct = RevitDuct.from_selection(uidoc, doc, view)
selected_duct = selected_duct[0] if selected_duct else None

if not selected_duct:
    output.print_md("## Select a duct first")
else:
    # Build the run from the selected duct
    run = RevitRuns.create_duct_run(selected_duct, doc, view)

    # Filter run to only straight/spiral ducts
    filtered_run = RevitHanger.filter_run_by_family(run, ALLOWED_FAMILIES)

    if not filtered_run:
        output.print_md("## No straight or spiral ducts found in run")
    else:
        # Calculate hanger positions
        positions = RevitHanger.calculate_hanger_positions(
            filtered_run,
            spacing_inches=HANGER_SPACING_INCHES,
            end_offset_inches=END_OFFSET_INCHES
        )

        if not positions:
            output.print_md(
                "## Run too short for hangers (need at least 12 feet)")
        else:
            # Get hanger type
            hanger_type = RevitHanger.get_hanger_type_by_name(
                doc, RevitHanger.HALF_STRAP_FAMILY)
            hanger_template = RevitHanger.get_hanger_instance_by_name(
                doc, RevitHanger.HALF_STRAP_FAMILY)
            if hanger_type:
                output.print_md(
                    "**Hanger Type Found:** ID {}".format(hanger_type.Id)
                )
            if hanger_template:
                output.print_md(
                    "**Template Hanger Found:** ID {}".format(
                        hanger_template.Id)
                )
            else:
                output.print_md(
                    "**Template Hanger Found:** None (will try API create path)"
                )
            if not hanger_type:
                output.print_md(
                    "## Half Strap Hanger family not found in project")
            elif not hanger_template:
                output.print_md(
                    "## Place one Half Strap Hanger manually in the model first, then run again."
                )
            else:
                output.print_md("---")
                output.print_md(
                    "### Placing {} hangers on run".format(len(positions)))

                created_count = 0
                failed_count = 0

                with revit.Transaction("Place Hangers on Run"):
                    for i, pos in enumerate(positions, start=1):
                        duct_idx = pos.get('duct_index', 0)
                        if duct_idx >= len(filtered_run):
                            duct_idx = len(filtered_run) - 1

                        target_duct = filtered_run[duct_idx]

                        result = RevitHanger.create_hanger_at_position(
                            doc,
                            hanger_type,
                            target_duct,
                            pos,
                            lower_elevation_offset_inches=HANGER_BELOW_ELEVATION_INCHES,
                            hanger_template=hanger_template
                        )

                        hanger = None
                        error_msg = None
                        if isinstance(result, tuple):
                            hanger, error_msg = result
                        else:
                            hanger = result

                        if hanger:
                            output.print_md(
                                "### {} | Created hanger ID: {}".format(
                                    i, output.linkify(hanger.Id)
                                )
                            )
                            created_count += 1
                        else:
                            if error_msg:
                                output.print_md(
                                    "### {} | Failed: {}".format(i, error_msg))
                            else:
                                output.print_md(
                                    "### {} | Failed to create hanger".format(i))
                            failed_count += 1

                output.print_md("---")
                output.print_md("## Summary")
                output.print_md("- **Created**: {}".format(created_count))
                output.print_md("- **Failed**: {}".format(failed_count))

                if created_count > 0:
                    output.print_md(
                        "- **Spacing**: {} feet".format(HANGER_SPACING_INCHES / 12.0))
                    output.print_md(
                        "- **End offset**: {} inches".format(END_OFFSET_INCHES))

                print_disclaimer(output)
