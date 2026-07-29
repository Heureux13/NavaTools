# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import script
from pyrevit import revit
import math

from Autodesk.Revit.DB import (
    BuiltInParameter,
    BuiltInCategory,
    ElementTransformUtils,
    ElementTypeGroup,
    FilteredElementCollector,
    HorizontalTextAlignment,
    Line,
    LocationPoint,
    TextNote,
    TextNoteOptions,
    XYZ,
)

# Button info
# ======================================================================
__title__ = 'Section View LBS'
__doc__ = '''
For every view symbol in the active view:
1) Read _UMI_PYT_SetionWeight
2) Create a centered text note at the symbol center
'''

# Variables
# ======================================================================

output = script.get_output()
PARAMETER_NAMES = (
    '_UMI_PYT_WeightSection',
    '_UMI_PYT_SetionWeight',
    '_UMI_PYT_SectionWeight',
)


def _param_to_text(param):
    """Return a readable parameter value as text."""
    if not param:
        return None

    text_value = param.AsString()
    if text_value:
        return text_value

    value_text = param.AsValueString()
    if value_text:
        return value_text

    # Fallback for numeric/integer values with no display string.
    try:
        dbl = param.AsDouble()
        if dbl is not None:
            return str(dbl)
    except Exception:
        pass

    try:
        intval = param.AsInteger()
        if intval is not None:
            return str(intval)
    except Exception:
        pass

    return None


def _symbol_center(element, view):
    """Get a best-effort center point for an element in the active view."""
    location = element.Location
    if isinstance(location, LocationPoint):
        return location.Point

    bbox = element.get_BoundingBox(view)
    if bbox:
        return XYZ(
            (bbox.Min.X + bbox.Max.X) * 0.5,
            (bbox.Min.Y + bbox.Max.Y) * 0.5,
            (bbox.Min.Z + bbox.Max.Z) * 0.5,
        )

    return None


def _is_symbol_vertical(element, view):
    """Infer symbol orientation from its 2D extents in the active view."""
    bbox = element.get_BoundingBox(view)
    if not bbox:
        return False

    width = abs(bbox.Max.X - bbox.Min.X)
    height = abs(bbox.Max.Y - bbox.Min.Y)
    return height > width


def _find_parameter_on_element_or_type(doc, element, parameter_names):
    """Return first matching parameter found on element or its type."""
    if not element:
        return None

    for pname in parameter_names:
        param = element.LookupParameter(pname)
        if param:
            return param

    type_id = element.GetTypeId()
    if type_id and type_id.IntegerValue > 0:
        elem_type = doc.GetElement(type_id)
        if elem_type:
            for pname in parameter_names:
                param = elem_type.LookupParameter(pname)
                if param:
                    return param

    return None


def _get_section_weight_text(doc, symbol):
    """Resolve section weight text from symbol, type, or referenced view."""
    param = _find_parameter_on_element_or_type(doc, symbol, PARAMETER_NAMES)
    text = _param_to_text(param)
    if text:
        return text

    ref_view_id_param = symbol.get_Parameter(BuiltInParameter.VIEWER_VIEW_ID)
    if ref_view_id_param:
        ref_view_id = ref_view_id_param.AsElementId()
        if ref_view_id and ref_view_id.IntegerValue > 0:
            ref_view = doc.GetElement(ref_view_id)
            param = _find_parameter_on_element_or_type(doc, ref_view, PARAMETER_NAMES)
            text = _param_to_text(param)
            if text:
                return text

    return None


doc = revit.doc
view = doc.ActiveView

view_symbols = (
    FilteredElementCollector(doc, view.Id)
    .OfCategory(BuiltInCategory.OST_Viewers)
    .WhereElementIsNotElementType()
    .ToElements()
)

if not view_symbols:
    output.print_md('No view symbols found in the active view.')
    script.exit()

text_type_id = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType)
if not text_type_id or text_type_id.IntegerValue < 0:
    output.print_md('No valid TextNote type found in this project.')
    script.exit()

options = TextNoteOptions(text_type_id)
options.HorizontalAlignment = HorizontalTextAlignment.Center

created = 0
missing_param = 0
missing_location = 0
rotated = 0

with revit.Transaction('Add section weight notes to view symbols'):
    for symbol in view_symbols:
        text = _get_section_weight_text(doc, symbol)
        if not text:
            missing_param += 1
            continue

        point = _symbol_center(symbol, view)
        if not point:
            missing_location += 1
            continue

        note = TextNote.Create(doc, view.Id, point, text, options)

        if _is_symbol_vertical(symbol, view):
            axis = Line.CreateBound(point, XYZ(point.X, point.Y, point.Z + 1.0))
            ElementTransformUtils.RotateElement(doc, note.Id, axis, math.pi / 2.0)
            rotated += 1

        created += 1

output.print_md('View symbols found: **{}**'.format(len(view_symbols)))
output.print_md('Notes created: **{}**'.format(created))
output.print_md('Notes rotated (vertical symbols): **{}**'.format(rotated))
output.print_md('Skipped (missing/empty section weight parameter): **{}**'.format(missing_param))
output.print_md('Skipped (no center point found): **{}**'.format(missing_location))
