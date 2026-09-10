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
__title__ = 'Follow Pipe'
__doc__ = '''
Places markers on selected pipes.
'''

# Variables
# ======================================================================
output = script.get_output()
BY_FAMILY = False

output.print_md('line 34')


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


def create_pipe_points(pipes):
    """Create a BIMrx_Point instance at each pipe bottom point."""
    output.print_md('Searching for BIMrx_Point type...')
    point_symbol = get_bimrx_point_symbol()
    output.print_md('BIMrx_Point type found: {}'.format(point_symbol.Id))
    created_count = 0
    created_ids = List[ElementId]()
    transaction = Transaction(revit.doc, 'Create Pipe Bottom Points')
    transaction.Start()
    output.print_md('Transaction started.')

    if not point_symbol.IsActive:
        point_symbol.Activate()
        revit.doc.Regenerate()

    for pipe in pipes:
        output.print_md('Getting pipe level...')
        level = revit.doc.GetElement(pipe.LevelId)
        if not isinstance(level, Level):
            raise ValueError('The pipe does not have a valid reference level.')

        output.print_md('Calculating pipe points...')
        for point in get_point(pipe):
            level_relative_point = XYZ(
                point.X,
                point.Y,
                point.Z - level.Elevation,
            )
            output.print_md('Placing point at {}'.format(level_relative_point))
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
            output.print_md(
                'Target Z: {:.6f}; bounding-box top Z: {:.6f}; '
                'instance origin Z: {:.6f}'.format(
                    point.Z,
                    moved_bounding_box.Max.Z,
                    instance.Location.Point.Z,
                )
            )
            created_ids.Add(instance.Id)
            created_count += 1

    transaction.Commit()

    revit.uidoc.Selection.SetElementIds(created_ids)
    revit.uidoc.ShowElements(created_ids)
    return created_count, created_ids


selected_pipes = []

output.print_md('Before reading selection')
selected_ids = revit.uidoc.Selection.GetElementIds()
output.print_md('Selection IDs read: {}'.format(selected_ids.Count))

for element_id in selected_ids:
    element = revit.doc.GetElement(element_id)
    output.print_md('Selected type: {}'.format(element.GetType().FullName))

    if (isinstance(element, FabricationPart)
            and isinstance(element.Location, LocationCurve)):
        selected_pipes.append(element)


output.print_md('Selected pipes: {}'.format(len(selected_pipes)))
if not selected_pipes:
    raise ValueError('Select at least one MEP pipe before running Follow Pipe.')

output.print_md('Finding BIMrx_Point type...')
created_count, created_ids = create_pipe_points(selected_pipes)
output.print_md(
    'Created BIMrx_Point instances: {}'.format(
        ', '.join(str(element_id) for element_id in created_ids)
    )
)
