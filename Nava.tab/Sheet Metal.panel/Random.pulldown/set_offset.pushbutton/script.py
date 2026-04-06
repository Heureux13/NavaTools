# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory
from geometry.offsets import Offsets
from geometry.size import Size
from ducts.revit_xyz import RevitXYZ
from pyrevit import revit, script
from Autodesk.Revit.DB import StorageType

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


# Collect fabrication ductwork in the current view only
doc = revit.doc
current_view = revit.active_view
all_fittings = FilteredElementCollector(doc, current_view.Id)\
    .OfCategory(BuiltInCategory.OST_FabricationDuctwork)\
    .WhereElementIsNotElementType()\
    .ToElements()

family_list = {
    'offset',
    'oval reducer',
    'oval to round',
    'reducer',
    'square to ø',
    'transition',
    'cid330 - (radius 2-way offset)',
    'ogee',
}

parameters = {
    '_offset_center_h': 'center_horizontal',
    '_offset_center_v': 'center_vertical',
    '_offset_top': 'top',
    '_offset_bottom': 'bottom',
    '_offset_right': 'right',
    '_offset_left': 'left',
}

tag_paramger = {
    '_offset'
}


def should_add_tag_prefix(classification):
    """Check if classification contains vertical UP/DN with numbers.

    Historically used to decide adding T: prefix. Kept for
    compatibility, now simply signals when UP/DN digits are present so
    we can convert to TU/TD format.
    """
    import re
    # Check if it contains UP or DN followed by a number (no space)
    if re.search(r'\b(UP|DN)\s*\d+', classification):
        return True
    return False


def convert_to_TU_TD(value):
    """Convert any 'T:UPn'/'T:DNn' or 'UP n'/'DN n' to 'TUn'/'TDn'.

    - Removes any leading 'T:' marker
    - Replaces vertical tokens:
        'UP12' or 'UP 12' -> 'TU12'
        'DN6'  or 'DN 6'  -> 'TD6'
    - Leaves other tokens (CL, FOB, FOT, IN, OUT, FOR, FOL) unchanged
    - Works within combined strings like 'UP12|IN5'
    """
    try:
        import re
        if value is None:
            return value
        s = value.strip()
        # drop optional leading T:
        if s.startswith('T:'):
            s = s[2:]
        # normalize spaces around vertical tokens and convert, append arrows
        s = re.sub(r'\bUP\s*(\d+)\b', r'TU\1', s)
        s = re.sub(r'\bDN\s*(\d+)\b', r'TD\1', s)
        # ensure arrows for TU/TD exactly once
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


output.print_md("# Offset Information")

# Filter by family name
matched_count = 0
processed = []
for element in all_fittings:
    try:
        # Get family name and normalize (trim spaces/asterisks) so *Reducer * matches 'reducer'
        family_name = doc.GetElement(element.GetTypeId()).FamilyName.lower()
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

        # Re-extract inlet/outlet data with size matching
        inlet_data, outlet_data = xyz_extractor.inlet_outlet_data(size)

        if not inlet_data or not outlet_data:
            continue

        # Calculate offsets with new API
        offsets_calc = Offsets(inlet_data, outlet_data, size)
        fitting = offsets_calc.calculate()

        if fitting:
            # Store for output
            processed.append((element, family_name, size_str,
                             fitting, inlet_data, outlet_data, size))

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
                                    p.Set(str(round(fitting[fitting_key], 3)))
                            except Exception:
                                pass

                # Write classification to tag offset parameter
                tag_p = element.LookupParameter('_offset')
                if tag_p and not tag_p.IsReadOnly:
                    try:
                        if tag_p.StorageType == StorageType.String:
                            current_value = tag_p.AsString()

                            # Convert the new classification to TU/TD scheme
                            final_classification = convert_to_TU_TD(
                                classification)

                            # If parameter has a value, replace only numbers while
                            # preserving existing format (user may have edited text)
                            if current_value and current_value.strip():
                                import re
                                result = current_value
                                # Extract all numbers from the new classification
                                numbers = re.findall(
                                    r'\d+', final_classification or '')
                                if numbers:
                                    for number in numbers:
                                        result = re.sub(
                                            r'\d+', number, result, count=1)
                                    tag_p.Set(result)
                                else:
                                    # No numbers to update, keep current value
                                    tag_p.Set(current_value)
                            else:
                                # Parameter is empty, write converted classification
                                tag_p.Set(final_classification)
                    except Exception:
                        pass

    except Exception:
        pass

# Print results
for i, (elem, fam, size, fit, inlet_data, outlet_data, size_obj) in enumerate(processed, start=1):
    cv = fit.get('center_vertical', 0)
    ch = fit.get('center_horizontal', 0)
    classification = convert_to_TU_TD(
        classify_offset(fit, inlet_data, outlet_data, size_obj))
    output.print_md(
        "### No: {:03} | ID: {} | Family: {} | Size: {} | CV: {:.2f}\" | CH: {:.2f}\" | {}".format(
            i,
            output.linkify(elem.Id),
            fam,
            size,
            cv,
            ch,
            classification
        ))

output.print_md("---")
output.print_md(
    "# Summary: {} fittings matched and processed".format(matched_count))
