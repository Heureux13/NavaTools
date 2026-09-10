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
    XYZ,
)
from Autodesk.Revit.DB.Structure import StructuralType
from pyrevit import script, revit
from System.Collections.Generic import List

# Button info
# ======================================================================
__title__ = 'All'
__doc__ = '''
Places markers on selected pipes.
'''

# Variables
# ======================================================================
output = script.get_output()
BY_FAMILY = False

DEBUG = False


def debug_print(message):
    """Print a message only when DEBUG is enabled."""
    if DEBUG:
        output.print_md(message)


ACCEPTED_FAMILIES = {
    'Pipe - PVC DWV Schedule 40 (PE x PE) - 20ft',
    'Pipe - CPVC Schedule 80 (PE x PE) - 20ft',
}


def get_fabrication_pipe_radius(pipe):
    """Return the physical outside radius from a round pipe connector."""
    for connector in pipe.ConnectorManager.Connectors:
        if connector.Radius > 0:
            return connector.Radius

    raise ValueError(
        'Could not find a round connector radius on fabrication part {}.'.format(
            pipe.Id,
        )
    )


def get_point(pipe):
    """Return the start, midpoint, and end XYZ locations at pipe bottom."""
    location = pipe.Location
    if not isinstance(location, LocationCurve):
        raise ValueError('The selected element does not have a location curve.')

    curve = location.Curve
    radius = get_fabrication_pipe_radius(pipe)
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
        for point in get_point(pipe):
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

            ElementTransformUtils.MoveElement(
                revit.doc,
                instance.Id,
                XYZ(0, 0, point.Z - bounding_box.Max.Z),
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

debug_print('Collecting fabrication pipes from active view...')
view_pipes = list(
    FilteredElementCollector(revit.doc, revit.active_view.Id)
    .OfClass(FabricationPart)
    .WhereElementIsNotElementType()
)
debug_print('Fabrication parts found in view: {}'.format(len(view_pipes)))

for element in view_pipes:
    if isinstance(element.Location, LocationCurve):
        element_type = revit.doc.GetElement(element.GetTypeId())
        if element_type.FamilyName in ACCEPTED_FAMILIES:
            selected_pipes.append(element)


debug_print('Selected pipes: {}'.format(len(selected_pipes)))
if not selected_pipes:
    raise ValueError('No matching MEP pipes found in the active view.')

debug_print('Finding BIMrx_Point type...')
created_count, created_ids = create_pipe_points(selected_pipes)
debug_print(
    'Created BIMrx_Point instances: {}'.format(
        ', '.join(str(element_id) for element_id in created_ids)
    )
)
