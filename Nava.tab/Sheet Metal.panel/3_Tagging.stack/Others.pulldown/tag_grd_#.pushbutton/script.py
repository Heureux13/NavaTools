# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
import re

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
Numbers all air terminals in the current view by level and scope box.
Writes numbers to the _# parameter in format like 2A001.
"""

# Helpers
# ==================================================


def _get_parameter_text(param, doc):
    if not param:
        return ""
    try:
        if param.StorageType == 4:  # ElementId
            elem_id = param.AsElementId()
            if elem_id and elem_id.IntegerValue > 0:
                linked_elem = doc.GetElement(elem_id)
                if linked_elem and getattr(linked_elem, "Name", None):
                    return linked_elem.Name
        val = param.AsString()
        if val is None:
            val = param.AsValueString()
        return val.strip() if val else ""
    except Exception:
        return ""


def _get_named_parameter_text_ci(element, doc, candidate_names):
    names = {n.lower().strip() for n in candidate_names}

    try:
        for param in element.Parameters:
            if not param or not param.Definition or not param.Definition.Name:
                continue
            if param.Definition.Name.lower().strip() in names:
                val = _get_parameter_text(param, doc)
                if val:
                    return val
    except Exception:
        pass

    try:
        elem_type = doc.GetElement(element.GetTypeId())
        if elem_type:
            for param in elem_type.Parameters:
                if not param or not param.Definition or not param.Definition.Name:
                    continue
                if param.Definition.Name.lower().strip() in names:
                    val = _get_parameter_text(param, doc)
                    if val:
                        return val
    except Exception:
        pass

    return ""


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


def _extract_scope_token(scope_name):
    """Extract the letter (A, B, C, etc.) from a scope box name like 'Area A', 'Area B'."""
    if not scope_name:
        return 'X'

    scope_up = scope_name.upper().strip()

    # Look for patterns like "AREA A", "BOX B", etc.
    match = re.search(r'[A-Z](?:\s|$)', scope_up)
    if match:
        return match.group(0).strip()

    # Fallback: get last alphanumeric token
    tokens = re.findall(r'[A-Z0-9]+', scope_up)
    if tokens:
        return tokens[-1]

    return 'X'


def _element_tag_point(element, active_view):
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


def _get_active_view_scope_name(active_view, doc):
    try:
        scope_param = active_view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)
        if scope_param and scope_param.StorageType == 4:
            scope_id = scope_param.AsElementId()
            if scope_id and scope_id.IntegerValue > 0:
                scope_elem = doc.GetElement(scope_id)
                if scope_elem and getattr(scope_elem, 'Name', None):
                    return scope_elem.Name
    except Exception:
        pass
    return ""


def _build_grd_numbers(elements, doc, active_view, manual_scope_letter=''):
    """Build GRD numbers grouped by level and scope."""
    all_levels = _get_all_levels_ordered(doc)
    scope_name = _get_active_view_scope_name(active_view, doc)

    # Use manual override if view scope is empty
    if not scope_name and manual_scope_letter:
        scope_token = manual_scope_letter.upper()
    else:
        scope_token = _extract_scope_token(scope_name)

    # DEBUG
    print("DEBUG: All Levels ({} found):".format(len(all_levels)))
    for idx, level in enumerate(all_levels, 1):
        try:
            print("  {}: {} (ID: {}, Elev: {})".format(idx, level.Name, level.Id.IntegerValue, level.Elevation))
        except Exception as e:
            print("  Error reading level: {}".format(e))
    print("DEBUG: Active View: {}".format(active_view.Name if active_view else "None"))
    print("DEBUG: Scope Name from view: '{}'".format(scope_name))
    print("DEBUG: Scope Token: '{}'".format(scope_token))
    grouped = {}

    for elem in elements:
        level_token = _extract_level_token(elem, doc, active_view, all_levels)
        group_key = "{}{}".format(level_token, scope_token)

        # DEBUG: Log first element for inspection
        if not grouped:
            print("DEBUG: First element (ID {}):".format(elem.Id.IntegerValue))
            try:
                elem_level_id = elem.LevelId
                print(
                    "  Element LevelId: {}".format(
                        elem_level_id.IntegerValue if elem_level_id and elem_level_id.IntegerValue > 0 else "Invalid"))
            except Exception as e:
                print("  Error reading LevelId: {}".format(e))
            print("  Level Token: '{}'".format(level_token))
            if active_view:
                try:
                    view_level = active_view.GenLevel
                    print(
                        "  View GenLevel: {}".format(
                            view_level.IntegerValue if view_level and view_level.IntegerValue > 0 else "Invalid"))
                except Exception as e:
                    print("  Error reading View GenLevel: {}".format(e))

        pt = _element_tag_point(elem, active_view)
        if pt:
            sort_key = (-round(pt.Y, 6), round(pt.X, 6), round(pt.Z, 6), elem.Id.IntegerValue)
        else:
            sort_key = (0.0, 0.0, 0.0, elem.Id.IntegerValue)

        grouped.setdefault(group_key, []).append((sort_key, elem))

    numbers_by_id = {}
    for group_key in sorted(grouped.keys()):
        group_elems = sorted(grouped[group_key], key=lambda x: x[0])
        for idx, (_, elem) in enumerate(group_elems, 1):
            numbers_by_id[elem.Id.IntegerValue] = "{}{num:0{pad}d}".format(
                group_key,
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
# Manual scope override (leave empty to read from view scope box)
# Set this if your view doesn't have a scope box assigned
manual_scope_letter = ''


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
    numbers_by_id = _build_grd_numbers(air_terminals, doc, view, manual_scope_letter)

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
