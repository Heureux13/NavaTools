# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from geometry.offsets import Offsets
from geometry.size import Size
from ducts.revit_xyz import RevitXYZ
from pyrevit import revit, script
from Autodesk.Revit.DB import StorageType
from config.parameters_registry import *

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
PRINT_OUTPUT = False

family_list = {
    'offset',
    'gored elbow',
    'ogee',
    'oval reducer',
    'oval to round',
    'reducer',
    'square to ø',
    'transition',
    'cid330 - (radius 2-way offset)'
}

parameters = {
    PYT_OFFSET_CENTER_H: 'center_horizontal',
    PYT_OFFSET_CENTER_V: 'center_vertical',
    PYT_OFFSET_TOP: 'top',
    PYT_OFFSET_BOTTOM: 'bottom',
    PYT_OFFSET_RIGHT: 'right',
    PYT_OFFSET_LEFT: 'left',
}

tag_paramger = {
    PYT_OFFSET_VALUE
}


def should_add_tag_prefix(classification):
    """Legacy detector for vertical UP/DN with numbers.

    Kept for compatibility to detect vertical tokens; we no longer add
    'T:' but use it to decide TU/TD conversion.
    """
    import re
    if re.search(r'\b(UP|DN)\s*\d+', classification):
        return True
    return False


def convert_to_TU_TD(value):
    """Convert UP/DN tokens to TU/TD with arrows (TU↑ / TD↓).

    - Removes leading 'T:' if present
    - UP12/UP 12 -> TU12↑
    - DN6/DN 6 -> TD6↓
    - Leaves other tokens alone
    """
    try:
        import re
        if value is None:
            return value
        s = value.strip()
        if s.startswith('T:'):
            s = s[2:]
        s = re.sub(r'\bUP\s*(\d+)\b', r'TU\1', s)
        s = re.sub(r'\bDN\s*(\d+)\b', r'TD\1', s)
        s = re.sub(r'\bTU(\d+)(?:[↑↓])?', r'TU\1↑', s)
        s = re.sub(r'\bTD(\d+)(?:[↑↓])?', r'TD\1↓', s)
        return s
    except Exception:
        return value


def classify_offset(fit, inlet_data=None, outlet_data=None, size_obj=None):
    """Classify offset values - vertical | horizontal.

    Vertical: FOB, FOT, CL, UP, DN
    Horizontal: FOL, FOR, CL, IN, OUT

    If top and bottom are both 0, just return horizontal.
    If left and right are both 0, just return vertical.
    If both are CL, return just CL.

    Args:
        fit: Dictionary with offset values
        inlet_data: Optional inlet connector data with basis vectors to normalize orientation
        outlet_data: Optional outlet connector data for world coordinates
    """
    tol = 0.01

    cv = fit.get('center_vertical', 0)
    ch = fit.get('center_horizontal', 0)
    top = fit.get('top', 0)
    bottom = fit.get('bottom', 0)
    right = fit.get('right', 0)
    left = fit.get('left', 0)

    rotated = False
    vertical_axis_sign = 1
    flow_sign = 1
    if inlet_data and inlet_data.get('basis_x') and inlet_data.get('basis_y'):
        bx = inlet_data['basis_x']
        by = inlet_data['basis_y']
        rotated = abs(bx.Z) > abs(by.Z) and abs(bx.Z) > 0.5
        vertical_axis_sign = 1 if (bx.Z if rotated else by.Z) >= 0 else -1
    if inlet_data and inlet_data.get('basis_z'):
        bz = inlet_data['basis_z']
        if abs(bz.X) >= abs(bz.Y):
            flow_sign = 1 if bz.X >= 0 else -1
        else:
            flow_sign = 1 if bz.Y >= 0 else -1

    def classify_vertical_part():
        """Classify vertical component from raw fit using axis orientation."""
        top_aligned = abs(top) < tol
        bottom_aligned = abs(bottom) < tol

        # Edge-aligned cases
        if top_aligned and bottom_aligned:
            if size_obj and size_obj.in_size == size_obj.out_size and abs(ch) >= tol:
                return "FOT" if vertical_axis_sign > 0 else "FOB"
            return "CL"
        if top_aligned:
            return "FOB" if flow_sign > 0 else "FOT"
        if bottom_aligned:
            return "FOT" if flow_sign > 0 else "FOB"

        # Neither edge aligned: determine UP/DN
        if abs(cv) < tol:
            return "CL"

        # Same-sign non-edge cases are orientation dependent.
        if top * bottom > 0:
            sign_source = 0.0
            if inlet_data and inlet_data.get('basis_y'):
                sign_source = inlet_data['basis_y'].Z
            if abs(sign_source) < 0.001:
                if rotated:
                    by = inlet_data.get('basis_y') if inlet_data else None
                    if by and (abs(by.Y) > abs(by.X)):
                        direction = "UP" if by.Y > 0 else "DN"
                    elif by:
                        direction = "UP" if by.X > 0 else "DN"
                    else:
                        direction = "UP" if vertical_axis_sign < 0 else "DN"
                    magnitude = max(abs(top), abs(bottom))
                    return "{} {:.0f}".format(direction, magnitude)
                sign_source = float(vertical_axis_sign)

            if sign_source > 0:
                direction = "UP"
                magnitude = min(abs(top), abs(bottom))
            else:
                direction = "DN"
                magnitude = max(abs(top), abs(bottom))
        else:
            # Mixed-sign edges require flow-aware interpretation.
            # For one flow direction the near-edge value is the label,
            # for the opposite flow direction the far-edge value is used.
            if flow_sign > 0:
                direction = "DN"
                magnitude = min(abs(top), abs(bottom))
            else:
                direction = "UP"
                magnitude = max(abs(top), abs(bottom))

        return "{} {:.0f}".format(direction, magnitude)

    def classify_horizontal_part():
        """Classify horizontal component using rotated/non-rotated rules."""
        if abs(left) < tol or abs(right) < tol:
            if size_obj and size_obj.in_size == size_obj.out_size and abs(ch) >= tol:
                direction = "IN" if (rotated and ch > 0) or (
                    not rotated and ch < 0) else "OUT"
                magnitude = max(abs(left), abs(right))
                return "{}{:.0f}".format(direction, magnitude)
            return "FOS"
        if abs(ch) < tol:
            return "CL"

        if rotated:
            direction = "IN" if ch > 0 else "OUT"
        else:
            direction = "OUT" if ch > 0 else "IN"

        if left * right > 0:
            if direction == "OUT":
                magnitude = min(abs(left), abs(right))
            else:
                magnitude = max(abs(left), abs(right))
        else:
            magnitude = max(abs(left), abs(right))

        return "{}{:.0f}".format(direction, magnitude)

    # If top and bottom are both 0, it's just a horizontal offset
    if abs(top) < tol and abs(bottom) < tol:
        horizontal = classify_horizontal_part()
        vertical = classify_vertical_part()
        if vertical == "CL":
            return horizontal
        return "{}|{}".format(vertical, horizontal)

    # If left and right are both 0, it's just a vertical offset
    elif abs(left) < tol and abs(right) < tol:
        vertical = classify_vertical_part()
        return vertical

    # Both vertical and horizontal
    else:
        vertical = classify_vertical_part()
        horizontal = classify_horizontal_part()

        if vertical == "CL" and horizontal == "CL":
            return "CL"

        return "{}|{}".format(vertical, horizontal)

# Main Code
# ======================================================================


# Get selected elements
selection = revit.get_selection()
doc = revit.doc

if not selection:
    if PRINT_OUTPUT:
        output.print_md(
            "# Please select offset fittings and try again"
        )

else:
    if PRINT_OUTPUT:
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
                processed.append(
                    (element, family_name, size_str, fitting, inlet_data, outlet_data))

                # Calculate classification
                classification = classify_offset(
                    fitting, inlet_data, outlet_data, size)

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
                    tag_p = element.LookupParameter('_offset')
                    if tag_p and not tag_p.IsReadOnly:
                        try:
                            if tag_p.StorageType == StorageType.String:
                                current_value = tag_p.AsString()
                                final_classification = convert_to_TU_TD(
                                    classification)
                                if current_value and current_value.strip():
                                    import re
                                    result = current_value
                                    numbers = re.findall(
                                        r'\d+', final_classification or '')
                                    if numbers:
                                        for number in numbers:
                                            result = re.sub(
                                                r'\d+', number, result, count=1)
                                        tag_p.Set(result)
                                    else:
                                        tag_p.Set(current_value)
                                else:
                                    tag_p.Set(final_classification)
                        except Exception:
                            pass

        except Exception as e:
            if PRINT_OUTPUT:
                output.print_md("ERROR: processing element {} : {}".format(
                    element.Id.Value,
                    str(e)
                ))

    # Print detailed results
    if PRINT_OUTPUT:
        for i, (elem, fam, size_str, fit, inlet_data, outlet_data) in enumerate(processed, start=1):
            size = Size(size_str)
            inlet = inlet_data['origin']
            outlet = outlet_data['origin']
            classification = convert_to_TU_TD(
                classify_offset(fit, inlet_data, outlet_data, size))

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

            # Add _duct_note parameter
            note_param = elem.LookupParameter('_duct_note')
            if note_param:
                note_value = note_param.AsString()
                output.print_md("### _duct_note | '{}'".format(note_value))

            # Add classification
            output.print_md("## Classification")
            output.print_md("### {}".format(classification))

        output.print_md("---")
        output.print_md(
            "# Summary: {} offset fittings processed, {} parameters updated".format(
                matched_count, len(processed)))
