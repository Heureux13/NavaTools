# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from offsets import Offsets
from size import Size
from revit_xyz import RevitXYZ
from pyrevit import revit, script
import traceback

# Button info
# ======================================================================
__title__ = 'Offset Data'
__doc__ = '''
Gives raw offset data
---
Values are representative of raw vector formulas used in linear mathmatics, seeing (-) and (+) does not mean in/out, it is movement from vector origins
'''

# Variables
# ======================================================================

output = script.get_output()

# Main Code
# ======================================================================

# Get selected elements
selection = revit.get_selection()

if not selection:
    output.print_md(
        "# Please select an offset and try again"
    )

else:
    output.print_md(
        "# Offset Information"
    )

    for element in selection:
        try:
            # Extract XYZ coordinates and orientation using RevitXYZ
            xyz_extractor = RevitXYZ(element)
            inlet_data, outlet_data = xyz_extractor.inlet_outlet_data()

            if not inlet_data or not outlet_data:
                output.print_md("ERROR: Element {} has no connector data".format(
                    element.Id.Value))
                continue

            inlet = inlet_data['origin']
            outlet = outlet_data['origin']

            # Parse size from element parameter
            size_param = element.LookupParameter("Size")
            if not size_param:
                output.print_md("ERROR: Element {} has no Size parameter".format(
                    element.Id.Value))
                continue

            size_str = size_param.AsString()
            size = Size(size_str)

            # Calculate offsets with new API
            offsets_calc = Offsets(inlet_data, outlet_data, size)
            fitting = offsets_calc.calculate()

            output.print_md("# Element ID: {} | Category {}".format(
                element.Id.Value,
                element.Category.Name))

            # General fitting information
            output.print_md(
                "### Size: {} | Inlet: {} | Outlet {}".format(
                    size.size,
                    size.in_size,
                    size.out_size,
                ))
            output.print_md(
                "### Inlet: {} | Shape: {}".format(
                    size.in_size,
                    size.in_shape(),
                ))
            output.print_md(
                "### Outlet: {} | Shape: {}".format(
                    size.out_size,
                    size.out_shape(),
                ))

            # Coordinate
            output.print_md("## **Coordinates**")
            output.print_md(
                "### Inlet: X: {:.3f}', Y: {:.3f}', Z: {:.3f}'".format(
                    inlet.X,
                    inlet.Y,
                    inlet.Z,
                ))
            output.print_md(
                "### Outlet: X: {:.3f}', Y: {:.3f}', Z: {:.3f}'".format(
                    outlet.X,
                    outlet.Y,
                    outlet.Z,
                ))

            if fitting:
                output.print_md("## Offset data")
                order = [
                    "center_vertical",
                    "center_horizontal",
                    "top",
                    "bottom",
                    "right",
                    "left",
                ]

                for key in order:
                    if key in fitting:
                        output.print_md("### {} | '{:.3f}'".format(
                            key,
                            fitting[key],
                        ))

            else:
                output.print_md("**ERROR**")
                output.print_md("Could not calculate for {}".format(
                    element.Id.Value
                ))

        except Exception as e:
            output.print_md("ERROR: processing element {} : {}".format(
                element.Id.Value,
                str(e)
            ))
            output.print_md("\n{}\n".format(
                traceback.format_exc()
            ))
