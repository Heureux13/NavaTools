# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory
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
    '_duct_offset_center_h': 'center_horizontal',
    '_duct_offset_center_v': 'center_vertical',
    '_duct_offset_top': 'top',
    '_duct_offset_bottom': 'bottom',
    '_duct_offset_right': 'right',
    '_duct_offset_left': 'left',
}

tag_paramger = {
    '_duct_tag_offset'
}


def should_add_tag_prefix(classification):
    """Check if classification should have T: prefix.

    Add T: if it contains UP or DN with numbers.
    Don't add T: for CL, FOB, FOT, FOR, FOL or combinations of these.
    """
    import re
    # Check if it contains UP or DN followed by a number (no space)
    if re.search(r'\b(UP|DN)\d+', classification):
        return True
    return False


def classify_offset(fit):
    """Classify offset values - vertical | horizontal.

    Vertical: FOB, FOT, CL, UP, DN
    Horizontal: FOL, FOR, CL, IN, OUT

    If top and bottom are both 0, just return horizontal.
    If left and right are both 0, just return vertical.
    If both are CL, return just CL.
    """
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
            return "{}{:.0f}".format(direction, magnitude)

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
            return "{}{:.0f}".format(direction, magnitude)

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

        # HORIZONTAL classification
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

        # Simplify if both are CL
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
                                    p.Set(str(round(fitting[fitting_key], 3)))
                            except Exception:
                                pass

                # Write classification to tag offset parameter
                tag_p = element.LookupParameter('_duct_tag_offset')
                if tag_p and not tag_p.IsReadOnly:
                    try:
                        if tag_p.StorageType == StorageType.String:
                            current_value = tag_p.AsString()

                            # Add T: prefix if needed
                            final_classification = classification
                            if should_add_tag_prefix(classification):
                                final_classification = "T:" + classification

                            # If parameter has a value, extract and replace only the numbers
                            if current_value and current_value.strip():
                                import re
                                # Extract all numbers from the new classification
                                numbers = re.findall(r'\d+', final_classification)
                                if numbers:
                                    # Replace numbers in existing value
                                    result = current_value
                                    # Add T: prefix if needed and not already there
                                    if should_add_tag_prefix(classification) and not result.startswith('T:'):
                                        result = 'T:' + result
                                    for number in numbers:
                                        result = re.sub(r'\d+', number, result, count=1)
                                    tag_p.Set(result)
                                else:
                                    # No numbers in classification, keep current value
                                    pass
                            else:
                                # Parameter is empty, write full classification
                                tag_p.Set(final_classification)
                    except Exception:
                        pass

    except Exception:
        pass

# Print results
for i, (elem, fam, size, fit) in enumerate(processed, start=1):
    cv = fit.get('center_vertical', 0)
    ch = fit.get('center_horizontal', 0)
    classification = classify_offset(fit)
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
