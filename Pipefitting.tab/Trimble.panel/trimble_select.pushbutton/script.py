# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.DB import (
    BuiltInCategory,
    ElementId,
    ElementTransformUtils,
    FamilyInstance,
    FamilySymbol,
    FabricationPart,
    FilteredElementCollector,
    Level,
    LocationCurve,
    Transaction,
    UnitTypeId,
    UnitUtils,
    XYZ,
)
from Autodesk.Revit.DB.Structure import StructuralType
from pyrevit import script, revit
import math
import traceback
from System.Collections.Generic import List
from pipefitting.sizes.pvc_sizes import SCHEDULE_40, SCHEDULE_80

# Button info
# ======================================================================
__title__ = 'Select'
__doc__ = '''
Places markers on selected pipes.
'''

# Variables
# ======================================================================
output = script.get_output()

BY_FAMILY = False
DEBUG = True
ACCEPTED_FAMILIES = {
    'Pipe - PVC DWV Schedule 40 (PE x PE) - 20ft': 'schedule_40',
    'Pipe - CPVC Schedule 80 (PE x PE) - 20ft': 'schedule_80',
}

schedule_lookup = {
    'schedule_40': SCHEDULE_40,
    'schedule_80': SCHEDULE_80,
}


def get_od_radius_pipe(pipe, element_type):
    schedule_key = ACCEPTED_FAMILIES[element_type.FamilyName]
    schedule = schedule_lookup[schedule_key]

    nominal_diameter = round(get_fabrication_pipe_radius(pipe) * 2, 2)
    size_data = schedule.get(nominal_diameter)
    if size_data is None:
        raise ValueError(
            "No chat entry for nominal size {} in {}".format(nominal_diameter, schedule_key)
        )
    od_inches = size_data['od']
    return ((od_inches) / 12.0) / 2.0


def debug_print(message):
    """Print a message only when DEBUG is enabled."""
    if DEBUG:
        output.print_md(message)


def get_param(element, name, unit=None, as_type="string", required=False):
    """Look up a parameter value on an element by name."""
    p = element.LookupParameter(name)
    if not p:
        if required:
            raise KeyError(
                "Missing parameter '{}' on element {}".format(
                    name,
                    element.Id,
                ))
        return None
    try:
        if as_type == "double":
            val = p.AsDouble()
            if val is None:
                return None
            if unit:
                val = UnitUtils.ConvertFromInternalUnits(val, unit)
            return float(val)
        if as_type == "int":
            return p.AsInteger()
        if as_type == "elementid":
            eid = p.AsElementId()
            return eid if isinstance(eid, ElementId) else None
        # fallback string: prefer AsString, then AsValueString
        s = p.AsString()
        if s is None:
            s = p.AsValueString()
        return s
    except Exception:
        # convert any unexpected Revit exception into None to keep callers
        # deterministic
        return None


def parse_inch_size(size_str):
    """Parse a Revit size string (e.g. '16"', '1 1/2"', '1-1/2"') into a float."""
    text = size_str.replace('"', '').strip()
    text = text.replace('-', ' ')
    parts = text.split()

    if not parts:
        raise ValueError("Could not parse size value '{}'.".format(size_str))

    whole = 0.0
    fraction = 0.0
    for part in parts:
        if '/' in part:
            numerator, denominator = part.split('/')
            fraction = float(numerator) / float(denominator)
        else:
            whole = float(part)

    return whole + fraction


def get_fabrication_pipe_radius(pipe):
    """Return the physical outside radius from a round pipe connector."""
    size_str = get_param(pipe, 'Size', as_type="string", required=True)
    if size_str is None:
        raise ValueError(
            'Could not find a round connector radius on fabrication part {}.'.format(
                pipe.Id,
            )
        )
    nominal_diameter = parse_inch_size(size_str)
    return nominal_diameter / 2


def get_pipe_slope_degrees(pipe):
    """Return the pipe's slope in degrees from horizontal, or None if unavailable."""
    p = pipe.LookupParameter('Slope')
    if not p or not p.HasValue:
        return None
    try:
        slope_ratio = p.AsDouble()
    except Exception:
        return None
    if slope_ratio is None:
        return None
    return math.degrees(math.atan(abs(slope_ratio)))


def get_point(pipe):
    """Return the start, midpoint, and end XYZ locations at pipe bottom."""
    location = pipe.Location
    if not isinstance(location, LocationCurve):
        raise ValueError('The selected element does not have a location curve.')
    element_type = revit.doc.GetElement(pipe.GetTypeId())

    curve = location.Curve
    radius = get_od_radius_pipe(pipe, element_type)
    direction = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize()
    vertical = XYZ.BasisZ
    bottom_direction = -vertical + direction.Multiply(direction.DotProduct(vertical))
    if bottom_direction.IsZeroLength():
        raise ValueError('Bottom points are undefined for a vertical pipe.')

    bottom_offset = bottom_direction.Normalize().Multiply(radius)

    return (
        curve.GetEndPoint(0) + bottom_offset,
        curve.Evaluate(0.5, True) + bottom_offset,
        curve.GetEndPoint(1) + bottom_offset,
    )


def get_bimrx_point_symbol():
    """Return the loaded BIMrx_Point family type."""
    symbols = (FilteredElementCollector(revit.doc)
               .OfClass(FamilySymbol)
               .OfCategory(BuiltInCategory.OST_GenericModel))

    for symbol in symbols:
        if symbol.Family.Name == 'BIMrx_Point':
            return symbol

    raise ValueError('The BIMrx_Point family is not loaded in this model.')


def get_existing_point_locations():
    """Return XYZ locations of all existing BIMrx_Point instances in the model."""
    instances = (FilteredElementCollector(revit.doc)
                 .OfClass(FamilyInstance)
                 .OfCategory(BuiltInCategory.OST_GenericModel))

    locations = []
    for instance in instances:
        if instance.Symbol.Family.Name == 'BIMrx_Point':
            bounding_box = instance.get_BoundingBox(None)
            origin = instance.Location.Point
            if bounding_box is None:
                locations.append(origin)
            else:
                locations.append(XYZ(origin.X, origin.Y, bounding_box.Max.Z))

    return locations


def point_already_exists(target_point, existing_locations, tolerance=0.01):
    """Check whether a point already exists near the target location."""
    for existing_point in existing_locations:
        if existing_point.DistanceTo(target_point) <= tolerance:
            return True
    return False


def create_pipe_points(pipes):
    """Create a BIMrx_Point instance at each pipe bottom point."""
    debug_print('Searching for BIMrx_Point type...')
    point_symbol = get_bimrx_point_symbol()
    debug_print('BIMrx_Point type found: {}'.format(point_symbol.Id))
    created_count = 0
    created_ids = List[ElementId]()
    existing_locations = get_existing_point_locations()
    transaction = Transaction(revit.doc, 'Create Pipe Bottom Points')
    transaction.Start()
    debug_print('Transaction started.')

    if not point_symbol.IsActive:
        point_symbol.Activate()
        revit.doc.Regenerate()

    for pipe in pipes:
        debug_print('Getting pipe level...')
        level = revit.doc.GetElement(pipe.LevelId)
        if not isinstance(level, Level):
            raise ValueError('The pipe does not have a valid reference level.')

        debug_print('Calculating pipe points...')
        try:
            pipe_points = get_point(pipe)
        except Exception:
            output.print_md('```\n{}\n```'.format(traceback.format_exc()))
            raise

        for point in pipe_points:
            if point_already_exists(point, existing_locations):
                debug_print(
                    'Skipping point at {}: already exists.'.format(point)
                )
                continue

            level_relative_point = XYZ(
                point.X,
                point.Y,
                point.Z - level.Elevation,
            )
            debug_print('Placing point at {}'.format(level_relative_point))
            instance = revit.doc.Create.NewFamilyInstance(
                level_relative_point,
                point_symbol,
                level,
                StructuralType.NonStructural,
            )
            revit.doc.Regenerate()
            bounding_box = instance.get_BoundingBox(None)
            if bounding_box is None:
                raise ValueError(
                    'Could not determine the BIMrx_Point family geometry bounds.'
                )

            bounding_box_center_z = (bounding_box.Max.Z + bounding_box.Min.Z) / 2.0
            ElementTransformUtils.MoveElement(
                revit.doc,
                instance.Id,
                XYZ(0, 0, point.Z - bounding_box_center_z),
            )
            revit.doc.Regenerate()
            moved_bounding_box = instance.get_BoundingBox(None)
            debug_print(
                'Target Z: {:.6f}; bounding-box top Z: {:.6f}; '
                'instance origin Z: {:.6f}'.format(
                    point.Z,
                    moved_bounding_box.Max.Z,
                    instance.Location.Point.Z,
                )
            )
            created_ids.Add(instance.Id)
            existing_locations.append(point)
            created_count += 1

    transaction.Commit()

    revit.uidoc.Selection.SetElementIds(created_ids)
    revit.uidoc.ShowElements(created_ids)
    return created_count, created_ids


selected_pipes = []

debug_print('Collecting fabrication pipes from current selection...')
selection_ids = revit.uidoc.Selection.GetElementIds()
selected_elements = [revit.doc.GetElement(eid) for eid in selection_ids]
selected_elements = [
    element for element in selected_elements
    if isinstance(element, FabricationPart)
]
debug_print('Fabrication parts found in selection: {}'.format(len(selected_elements)))

for element in selected_elements:
    if isinstance(element.Location, LocationCurve):
        element_type = revit.doc.GetElement(element.GetTypeId())
        if element_type.FamilyName in ACCEPTED_FAMILIES:
            slope_degrees = get_pipe_slope_degrees(element)
            if slope_degrees is None:
                debug_print(
                    'Excluding pipe {}: slope is not computed or missing.'.format(
                        element.Id,
                    )
                )
                continue
            if slope_degrees >= 45:
                debug_print(
                    'Excluding pipe {}: slope {:.2f} degrees is not under 45.'.format(
                        element.Id,
                        slope_degrees,
                    )
                )
                continue
            selected_pipes.append(element)


debug_print('Selected pipes: {}'.format(len(selected_pipes)))
if not selected_pipes:
    raise ValueError('No matching MEP pipes found in the current selection.')

debug_print('Finding BIMrx_Point type...')
created_count, created_ids = create_pipe_points(selected_pipes)
debug_print(
    'Created BIMrx_Point instances: {}'.format(
        ', '.join(str(element_id) for element_id in created_ids)
    )
)
