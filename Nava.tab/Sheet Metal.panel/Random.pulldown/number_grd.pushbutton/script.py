# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from pyrevit import revit, script, DB
from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    FilteredElementCollector,
    Transaction,
    XYZ,
)

# Button display information
# =================================================
__title__ = "Number GRDs"
__doc__ = """
Numbers all air terminals in the current view by level.
Writes numbers to the _# parameter in format like 1-0001.
"""

# Helpers
# ==================================================


def _get_all_levels_ordered(doc):
    """Get all levels in the project, ordered by elevation."""
    levels = []
    try:
        all_levels = FilteredElementCollector(doc).OfClass(DB.Level).ToElements()
        for level in all_levels:
            try:
                elev = level.Elevation
                levels.append((elev, level))
            except Exception:
                pass
    except Exception:
        pass
    levels.sort(key=lambda x: x[0])
    return [level for _, level in levels]


def _extract_level_token(element, doc, active_view, all_levels):
    """Get the sequential level number (1, 2, 3, etc.) for this element."""
    level_id = None

    # First try to get level from element
    try:
        elem_level_id = element.LevelId
        if elem_level_id and elem_level_id.IntegerValue > 0:
            level_id = elem_level_id
    except Exception:
        pass

    # If element has no level and we're in a plan view, get level from view
    if not level_id and active_view:
        try:
            view_level = active_view.GenLevel
            # GenLevel returns a Level object, not an ElementId
            if view_level:
                level_id = view_level.Id
        except Exception:
            pass

    # Match level ID to sequential index
    if level_id:
        try:
            for idx, level in enumerate(all_levels, 1):
                if level.Id.IntegerValue == level_id.IntegerValue:
                    return str(idx)
        except Exception:
            pass

    return 'X'


def _element_tag_point(element, active_view):
    """Get the center point of an element for sorting."""
    try:
        loc = element.Location
        if hasattr(loc, 'Point'):
            return loc.Point
    except Exception:
        pass

    try:
        bbox = element.get_BoundingBox(active_view) if active_view else None
        if bbox:
            min_pt = bbox.Min
            max_pt = bbox.Max
            return XYZ(
                (min_pt.X + max_pt.X) / 2.0,
                (min_pt.Y + max_pt.Y) / 2.0,
                (min_pt.Z + max_pt.Z) / 2.0,
            )
    except Exception:
        pass

    return None


def _build_grd_numbers(elements, doc, active_view):
    """Build GRD numbers grouped by level only, sorted by position."""
    all_levels = _get_all_levels_ordered(doc)

    print("DEBUG: All Levels ({} found):".format(len(all_levels)))
    for idx, level in enumerate(all_levels, 1):
        try:
            print("  {}: {} (Elev: {})".format(idx, level.Name, level.Elevation))
        except Exception as e:
            print("  Error reading level: {}".format(e))

    grouped = {}

    for elem in elements:
        level_token = _extract_level_token(elem, doc, active_view, all_levels)

        pt = _element_tag_point(elem, active_view)
        if pt:
            sort_key = (-round(pt.Y, 6), round(pt.X, 6), round(pt.Z, 6), elem.Id.IntegerValue)
        else:
            sort_key = (0.0, 0.0, 0.0, elem.Id.IntegerValue)

        grouped.setdefault(level_token, []).append((sort_key, elem))

    numbers_by_id = {}
    for level_token in sorted(grouped.keys()):
        group_elems = sorted(grouped[level_token], key=lambda x: x[0])
        for idx, (_, elem) in enumerate(group_elems, 1):
            numbers_by_id[elem.Id.IntegerValue] = "{}-{num:0{pad}d}".format(
                level_token,
                num=idx,
                pad=number_padding,
            )

    return numbers_by_id


# Code
# ==================================================
doc = revit.doc
view = revit.active_view
output = script.get_output()

# Numbering configuration
# ==================================================
number_parameter = '_#'
number_padding = 4

air_terminals = (
    FilteredElementCollector(doc, view.Id)
    .OfCategory(BuiltInCategory.OST_DuctTerminal)
    .WhereElementIsNotElementType()
    .ToElements()
)

if not air_terminals:
    script.exit()

t = Transaction(doc, "Number Air Terminals")
t.Start()
try:
    numbers_by_id = _build_grd_numbers(air_terminals, doc, view)

    # Number all air terminals
    for elem in air_terminals:
        grd_number = numbers_by_id.get(elem.Id.IntegerValue, "")

        try:
            param = elem.LookupParameter(number_parameter)
            if param:
                param.Set(grd_number)
        except Exception:
            pass

    t.Commit()
except Exception as e:
    t.RollBack()
    raise
