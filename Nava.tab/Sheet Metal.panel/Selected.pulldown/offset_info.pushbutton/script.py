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
from Autodesk.Revit.DB import StorageType
import traceback

# Button info
# ======================================================================
__title__ = 'Offset Data'
__doc__ = '''
Gives raw offset data and writes to parameters
---
Values are representative of raw vector formulas used in linear mathmatics, seeing (-) and (+) does not mean in/out, it is movement from vector origins
'''

# Variables
# ======================================================================

output = script.get_output()

family_list = {
    'offset',
    'gored elbow',
    'oval reducer',
    'oval to round',
    'reducer',
    'square to ø',
    'transition',
    'cid330 - (radius 2-way offset)'
}

parameters = {
    '_duct_offset_center_h': 'center_horizontal',
    '_duct_offset_center_v': 'center_vertical',
    '_duct_offset_top': 'top',
    '_duct_offset_bottom': 'bottom',
    '_duct_offset_right': 'right',
    '_duct_offset_left': 'left',
}

tag_parameter = {
    '_duct_tag_offset'
}


def classify_offset(fit):
    """Classify offset values - vertical | horizontal."""
    tol = 0.01

    cv = fit.get('center_vertical', 0)
    ch = fit.get('center_horizontal', 0)
    top = fit.get('top', 0)
    bottom = fit.get('bottom', 0)
    right = fit.get('right', 0)
    left = fit.get('left', 0)

    # If top and bottom are both 0, it's just a horizontal offset
    if abs(top) < tol and abs(bottom) < tol:
        if abs(left) < tol:
            return "FOR"
        elif abs(right) < tol:
            return "FOL"
        elif abs(ch) < tol:
            return "CL"
        else:
            magnitude = max(abs(left), abs(right))
            # Use the dominant edge sign to decide IN/OUT (positive right => OUT)
            if abs(right) >= abs(left):
                direction = "OUT" if right > 0 else "IN"
            else:
                direction = "IN" if left > 0 else "OUT"
            return "{} {:.0f}".format(direction, magnitude)

    # If left and right are both 0, it's just a vertical offset
    elif abs(left) < tol and abs(right) < tol:
        # Determine FOB/FOT based on center_vertical and edge alignment
        # FOB: outlet at bottom edge (cv<0 & bottom≈0) OR (cv>0 & top≈0)
        # FOT: outlet at top edge (cv<0 & top≈0) OR (cv>0 & bottom≈0)
        if abs(bottom) < tol or abs(top) < tol:
            # Check relationship between center_vertical and which edge is zero
            if (cv < -tol and abs(bottom) < tol) or (cv > tol and abs(top) < tol):
                return "FOB"
            elif (cv < -tol and abs(top) < tol) or (cv > tol and abs(bottom) < tol):
                return "FOT"
            # Fallback to original logic if cv is near zero
            elif abs(cv) < tol:
                return "FOB" if abs(bottom) < tol else "FOT"
            else:
                # Default case
                return "FOB" if abs(bottom) < tol else "FOT"
        elif abs(cv) < tol:
            return "CL"
        else:
            magnitude = max(abs(top), abs(bottom))
            # Flip vertical sense per user: positive center_v => DN
            direction = "DN" if cv > 0 else "UP"
            return "{} {:.0f}".format(direction, magnitude)

    # Both vertical and horizontal
    else:
        # Vertical classification with flow-direction correction
        if abs(bottom) < tol or abs(top) < tol:
            if (cv < -tol and abs(bottom) < tol) or (cv > tol and abs(top) < tol):
                vertical = "FOB"
            elif (cv < -tol and abs(top) < tol) or (cv > tol and abs(bottom) < tol):
                vertical = "FOT"
            elif abs(cv) < tol:
                vertical = "FOB" if abs(bottom) < tol else "FOT"
            else:
                vertical = "FOB" if abs(bottom) < tol else "FOT"
        elif abs(cv) < tol:
            vertical = "CL"
        else:
            magnitude = max(abs(top), abs(bottom))
            # Flip vertical sense per user: positive center_v => DN
            direction = "DN" if cv > 0 else "UP"
            vertical = "{} {:.0f}".format(direction, magnitude)

        if abs(left) < tol:
            horizontal = "FOR"
        elif abs(right) < tol:
            horizontal = "FOL"
        elif abs(ch) < tol:
            horizontal = "CL"
        else:
            magnitude = max(abs(left), abs(right))
            # Use the dominant edge sign to decide IN/OUT (positive right => OUT)
            if abs(right) >= abs(left):
                direction = "OUT" if right > 0 else "IN"
            else:
                direction = "IN" if left > 0 else "OUT"
            horizontal = "{} {:.0f}".format(direction, magnitude)

        if vertical == "CL" and horizontal == "CL":
            return "CL"

        return "{} | {}".format(vertical, horizontal)

# Main Code
# ======================================================================


# Get selected elements
selection = revit.get_selection()
doc = revit.doc

if not selection:
    output.print_md(
        "# Please select offset fittings and try again"
    )

else:
    output.print_md(
        "# Offset Information"
    )

    processed = []
    matched_count = 0

    for element in selection:
        try:
            # Get family name and normalize
            family_type = doc.GetElement(element.GetTypeId())
            if not family_type:
                continue

            family_name = family_type.FamilyName.lower()
            family_name = family_name.replace('*', '').strip()

            if family_name not in family_list:
                continue

            matched_count += 1

            # Extract XYZ coordinates and orientation using RevitXYZ
            xyz_extractor = RevitXYZ(element)
            inlet_data, outlet_data = xyz_extractor.inlet_outlet_data()

            if not inlet_data or not outlet_data:
                continue

            inlet = inlet_data['origin']
            outlet = outlet_data['origin']

            # Parse size from element parameter
            size_param = element.LookupParameter("Size")
            if not size_param:
                continue

            size_str = size_param.AsString()
            size = Size(size_str)

            # Extract XYZ with size matching - re-extract with size info
            inlet_data, outlet_data = xyz_extractor.inlet_outlet_data(size)

            # Calculate offsets with new API
            offsets_calc = Offsets(inlet_data, outlet_data, size)
            fitting = offsets_calc.calculate()

            if fitting:
                # Store for output
                processed.append((element, family_name, size_str, fitting))

                # Calculate classification
                classification = classify_offset(fitting)

                # Write values to parameters
                with revit.Transaction("Set Offset Parameters"):
                    for param_name, fitting_key in parameters.items():
                        if fitting_key in fitting:
                            p = element.LookupParameter(param_name)
                            if p and not p.IsReadOnly:
                                try:
                                    if p.StorageType == StorageType.Double:
                                        p.Set(fitting[fitting_key])
                                    elif p.StorageType == StorageType.String:
                                        p.Set(
                                            str(round(fitting[fitting_key], 3)))
                                except Exception:
                                    pass

                    # Write classification to tag offset parameter
                    tag_p = element.LookupParameter('_duct_tag_offset')
                    if tag_p and not tag_p.IsReadOnly:
                        try:
                            if tag_p.StorageType == StorageType.String:
                                tag_p.Set(classification)
                        except Exception:
                            pass

        except Exception as e:
            output.print_md("ERROR: processing element {} : {}".format(
                element.Id.Value,
                str(e)
            ))

    # Print detailed results
    for i, (elem, fam, size_str, fit) in enumerate(processed, start=1):
        size = Size(size_str)
        inlet_data, outlet_data = RevitXYZ(elem).inlet_outlet_data()
        inlet = inlet_data['origin']
        outlet = outlet_data['origin']
        classification = classify_offset(fit)

        output.print_md("# Element ID: {} | Category {}".format(
            elem.Id.Value,
            elem.Category.Name))

        # General fitting information
        output.print_md(
            "### Size: {} | Inlet: {} | Outlet {}".format(
                size.size,
                size.in_size,
                size.out_size,
            ))
        output.print_md(
            "### Inlet: {} | Shape: {} | W:{} H:{} D:{}".format(
                size.in_size,
                size.in_shape(),
                size.in_width,
                size.in_height,
                size.in_diameter,
            ))
        output.print_md(
            "### Outlet: {} | Shape: {} | W:{} H:{} D:{}".format(
                size.out_size,
                size.out_shape(),
                size.out_width,
                size.out_height,
                size.out_diameter,
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

        # Debug: Show basis vectors
        output.print_md("## **Basis Vectors (Debug)**")
        if inlet_data.get('basis_z'):
            bz = inlet_data['basis_z']
            output.print_md(
                "### Inlet Flow: X: {:.3f}, Y: {:.3f}, Z: {:.3f}".format(bz.X, bz.Y, bz.Z))
        if outlet_data.get('basis_z'):
            bz = outlet_data['basis_z']
            output.print_md(
                "### Outlet Flow: X: {:.3f}, Y: {:.3f}, Z: {:.3f}".format(bz.X, bz.Y, bz.Z))
        if inlet_data.get('basis_y'):
            by = inlet_data['basis_y']
            output.print_md(
                "### Inlet Up: X: {:.3f}, Y: {:.3f}, Z: {:.3f}".format(by.X, by.Y, by.Z))
        if inlet_data.get('basis_x'):
            bx = inlet_data['basis_x']
            output.print_md(
                "### Inlet Right: X: {:.3f}, Y: {:.3f}, Z: {:.3f}".format(bx.X, bx.Y, bx.Z))

        # Offset data
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
            if key in fit:
                output.print_md("### {} | '{:.3f}'".format(
                    key,
                    fit[key],
                ))

        # Add classification
        output.print_md("## Classification")
        output.print_md("### {}".format(classification))

    output.print_md("---")
    output.print_md(
        "# Summary: {} offset fittings processed, {} parameters updated".format(
            matched_count, len(processed)))
