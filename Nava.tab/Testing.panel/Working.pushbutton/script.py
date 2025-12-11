# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from offsets import Offsets
from pyrevit import revit, script

# Button info
# ===================================================
__title__ = "Test Offsets"
__doc__ = """
Test the Offsets class with selected duct/pipe elements
"""

# Variables
# ==================================================
output = script.get_output()

# Main
# ==================================================

# Get selected elements
selection = revit.get_selection()

if not selection:
    output.print_md("**Please select at least one duct or pipe element**")
else:
    output.print_md("# TESTING OFFSETS CLASS")

    for element in selection:
        try:
            # Try to create Offsets object
            offsets_obj = Offsets(element)

            # Check if we have valid data
            if offsets_obj.start_point and offsets_obj.end_point and offsets_obj.size:
                output.print_md("---")
                output.print_md("**Element ID:** {}".format(element.Id.Value))
                output.print_md(
                    "**Element Type:** {}".format(element.Category.Name))

                # Print size info
                if offsets_obj.size:
                    output.print_md(
                        "**Size:** {}".format(offsets_obj.size.size))
                    output.print_md(
                        "  - Inlet: {} ({})".format(offsets_obj.size.in_size, offsets_obj.size.in_shape()))
                    output.print_md(
                        "  - Outlet: {} ({})".format(offsets_obj.size.out_size, offsets_obj.size.out_shape()))

                # Print inlet and outlet coordinates (in feet)
                output.print_md("**Coordinates (feet):**")
                output.print_md("  - Inlet: X={:.3f}, Y={:.3f}, Z={:.3f}".format(
                    offsets_obj.start_point.X, offsets_obj.start_point.Y, offsets_obj.start_point.Z))
                output.print_md("  - Outlet: X={:.3f}, Y={:.3f}, Z={:.3f}".format(
                    offsets_obj.end_point.X, offsets_obj.end_point.Y, offsets_obj.end_point.Z))

                # Calculate offsets
                result = offsets_obj.calculate_offsets()

                if result:
                    output.print_md("**Calculated Offsets (inches):**")
                    for key, value in result.items():
                        output.print_md(
                            "  - **{}:** `{:.3f}`".format(key, value))
                else:
                    output.print_md(
                        "**ERROR:** Could not calculate offsets for element {}".format(element.Id.Value))
            else:
                output.print_md("---")
                output.print_md("**Element ID:** {}".format(element.Id.Value))
                output.print_md(
                    "**Element Type:** {}".format(element.Category.Name))

                # Detailed diagnostics
                output.print_md("**Diagnostics:**")
                output.print_md(
                    "  - Has start point: {}".format(offsets_obj.start_point is not None))
                output.print_md(
                    "  - Has end point: {}".format(offsets_obj.end_point is not None))
                output.print_md(
                    "  - Has size: {}".format(offsets_obj.size is not None))

                if offsets_obj.size:
                    output.print_md("  - Size string parsed successfully")
                    output.print_md(
                        "    - in_size: {}".format(offsets_obj.size.in_size))
                    output.print_md(
                        "    - out_size: {}".format(offsets_obj.size.out_size))

                error_msg = getattr(offsets_obj, 'error_msg',
                                    None) or "Unknown issue"
                output.print_md("**Reason:** {}".format(error_msg))

        except Exception as e:
            output.print_md(
                "**ERROR:** processing element {}: {}".format(element.Id.Value, str(e)))
            import traceback
            output.print_md("```\n{}\n```".format(traceback.format_exc()))
            print("ERROR processing element {}: {}".format(
                element.Id.Value, str(e)))
            import traceback
            print(traceback.format_exc())
