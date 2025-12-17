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


# Collect all fabrication ductwork in the document
doc = revit.doc
all_fittings = FilteredElementCollector(doc)\
    .OfCategory(BuiltInCategory.OST_FabricationDuctwork)\
    .WhereElementIsNotElementType()\
    .ToElements()

family_list = {
    'offset',
    'gored elbow',
    'oval reducer',
    'oval to round',
    'reducer',
    'square to ø',
    'transition',
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
            return "{} {:.0f}".format(direction, magnitude)

    # If left and right are both 0, it's just a vertical offset
    elif abs(left) < tol and abs(right) < tol:
        if abs(bottom) < tol:
            return "FOB"
        elif abs(top) < tol:
            return "FOT"
        elif abs(cv) < tol:
            return "CL"
        else:
            magnitude = max(abs(top), abs(bottom))
            # Flip vertical sense per user: positive center_v => DN
            direction = "DN" if cv > 0 else "UP"
            return "{} {:.0f}".format(direction, magnitude)

    # Both vertical and horizontal
    else:
        # VERTICAL classification
        if abs(bottom) < tol:
            vertical = "FOB"
        elif abs(top) < tol:
            vertical = "FOT"
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

        return "{} | {}".format(vertical, horizontal)


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
                            tag_p.Set(classification)
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
