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
    FamilySymbol,
    IndependentTag,
    Transaction,
    XYZ,
)
from revit_tagging import RevitTagging

# Button display information
# =================================================
__title__ = "Tag GRDs"
__doc__ = """
Tags all air terminals in the current view with the -UMI_GRD_JN label.
"""

# Helpers
# ==================================================


def _find_tag_symbol(doc, target_name):
    """Return the first air terminal tag whose name contains target_name."""
    if not target_name:
        return None
    needle = target_name.strip().lower()
    symbols = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_DuctTerminalTags)
        .OfClass(FamilySymbol)
        .ToElements()
    )
    exact_matches = []
    contains_matches = []
    for sym in symbols:
        fam = getattr(sym, "Family", None)
        fam_name = fam.Name if fam else ""
        type_name = getattr(sym, "Name", "") or ""
        fam_norm = fam_name.strip().lower()
        type_norm = type_name.strip().lower()
        label = (fam_name + " " + type_name).lower()
        if needle == fam_norm or needle == type_norm:
            exact_matches.append(sym)
        elif needle in label:
            contains_matches.append(sym)

    if exact_matches:
        return exact_matches[0]
    if contains_matches:
        return contains_matches[0]
    return None


def _tag_type_matches_target(tag_type, target_name):
    if not tag_type or not target_name:
        return False
    target = target_name.strip().lower()
    fam = getattr(tag_type, 'Family', None)
    fam_name = (fam.Name if fam else "").strip().lower()
    type_name = (getattr(tag_type, 'Name', '') or '').strip().lower()
    if target == fam_name or target == type_name:
        return True
    label = (fam_name + " " + type_name).strip()
    return target in label


# Code
# ==================================================
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

doc = revit.doc
view = revit.active_view
output = script.get_output()

tagger = RevitTagging(doc, view)

# Tag sets in priority order
first_tag = [
    '_umi_grd_cfm',
]

second_tag = [
    '_umi_grd',
]

# Parameters to check and their "empty" values
check_parameter = {
    'airflow': [0, "0", "0 cfm"],
    'flow': [0, "0", "0 cfm"],
    'cfm': [0, "0", "0 cfm"]
}

order_paramters = {
    1: '_grd_value',
    2: 'mark',
    3: 'type mark',
}

value_parameters = {
    '_grd_label',
}

skip_values = {
    'skip',
    'n/a',
}

second_tag_values = {
    'second',
    '2',
    '_umi_grd'
}


def _has_real_value(element, param_name, empty_values):
    """Check if a parameter has a non-empty value."""
    try:
        param = element.LookupParameter(param_name)
        if not param:
            return False

        # Get value based on storage type
        storage_type = param.StorageType
        if storage_type == 1:  # Integer
            val = param.AsInteger()
        elif storage_type == 2:  # Double
            val = param.AsDouble()
        elif storage_type == 3:  # String
            val = param.AsString()
        else:
            val = param.AsValueString()

        # Check if value is not in empty_values list
        if val not in empty_values:
            return True
    except Exception:
        pass
    return False


def _get_param_case_insensitive(element, param_name):
    target = param_name.lower().strip()
    for param in element.Parameters:
        try:
            if param.Definition and param.Definition.Name and param.Definition.Name.lower().strip() == target:
                return param
        except Exception:
            pass
    return None


def _is_empty_parameter_value(param, empty_values):
    if not param:
        return False

    empty_norm = {str(v).lower().strip() for v in empty_values}
    try:
        storage_type = param.StorageType
        if storage_type == 1:  # Integer
            val_int = param.AsInteger()
            if val_int == 0:
                return True
        elif storage_type == 2:  # Double
            val_dbl = param.AsDouble()
            if val_dbl is not None and abs(val_dbl) < 1e-9:
                return True

        val_str = param.AsString()
        if val_str is None:
            val_str = param.AsValueString()
        if val_str is not None:
            val_norm = val_str.lower().strip()
            if val_norm in empty_norm:
                return True

            # Handle formatted strings like "0.00 CFM" by parsing the leading number.
            try:
                first_token = val_norm.split()[0].replace(',', '')
                if float(first_token) == 0.0:
                    return True
            except Exception:
                pass
    except Exception:
        pass

    return False


def _find_first_available_tag(doc, tag_names):
    """Try to find the first available tag from a list of tag names."""
    for tag_name in tag_names:
        tag_sym = _find_tag_symbol(doc, tag_name)
        if tag_sym:
            return tag_sym, tag_name
    return None, None


def _get_value_from_ordered_params(element, doc, param_order):
    """Get value from element following the ordered parameter hierarchy.
    Returns the first non-empty value found, or empty string if none found.
    Does case-insensitive parameter lookup.
    """
    elem_type = doc.GetElement(element.GetTypeId())

    for order, param_name in param_order.items():
        param_name_lower = param_name.lower().strip()

        # Try instance parameters with case-insensitive lookup
        for param in element.Parameters:
            if param.Definition.Name.lower().strip() == param_name_lower:
                try:
                    val = param.AsString()
                    if not val:
                        val = param.AsValueString()
                    if val and val.strip():
                        return val.strip()
                except Exception:
                    pass

        # If not found on instance, try type parameters with case-insensitive lookup
        if elem_type:
            for param in elem_type.Parameters:
                if param.Definition.Name.lower().strip() == param_name_lower:
                    try:
                        val = param.AsString()
                        if not val:
                            val = param.AsValueString()
                        if val and val.strip():
                            return val.strip()
                    except Exception:
                        pass

    return ""


air_terminals = (
    FilteredElementCollector(doc, view.Id)
    .OfCategory(BuiltInCategory.OST_DuctTerminal)
    .WhereElementIsNotElementType()
    .ToElements()
)

if not air_terminals:
    # output.print_md("## No air terminals found in this view.")
    script.exit()

placed = []
failed = []
already_tagged = []

# Check how many tags already exist in the view
existing_tags = list(
    FilteredElementCollector(doc, view.Id)
    .OfClass(IndependentTag)
    .ToElements()
)

# Build a map of element IDs to existing tag instances from our families
tag_map = {}
all_tag_names = first_tag + second_tag

for tag in existing_tags:
    try:
        tag_type_id = tag.GetTypeId()
        tag_type = doc.GetElement(tag_type_id)
        if not tag_type:
            continue

        is_our_tag = False
        for tag_name in all_tag_names:
            if _tag_type_matches_target(tag_type, tag_name):
                is_our_tag = True
                break

        if not is_our_tag:
            continue

        try:
            tagged_ids = tag.GetTaggedLocalElementIds()
        except Exception:
            tagged_ids = []

        for tid in tagged_ids:
            tid_val = tid.IntegerValue if hasattr(tid, 'IntegerValue') else int(tid)
            tag_map.setdefault(tid_val, []).append(tag)
    except BaseException:
        pass

t = Transaction(doc, "Tag Air Terminals")
t.Start()
try:
    # Update value parameters for all air terminals before tagging
    for elem in air_terminals:
        # Get value based on ordered parameter hierarchy
        value_to_write = _get_value_from_ordered_params(elem, doc, order_paramters)

        # Write to all value parameters
        for param_name in value_parameters:
            try:
                param = elem.LookupParameter(param_name)
                if param:
                    param.Set(value_to_write)
            except Exception:
                pass

        # Get the value that will be written
        current_value = _get_value_from_ordered_params(elem, doc, order_paramters)
        current_value_normalized = current_value.lower().strip()

        # Check if value is in skip list
        if current_value_normalized in {v.lower().strip() for v in skip_values}:
            failed.append((elem, "Value '{}' is in skip list".format(current_value)))
            continue

        # Check airflow/cfm parameters to determine which tag to use
        # If airflow is empty (0, "0", "0 cfm"), use second_tag; otherwise use first_tag
        use_second = False
        for param_name, empty_vals in check_parameter.items():
            param = _get_param_case_insensitive(elem, param_name)
            if param and _is_empty_parameter_value(param, empty_vals):
                use_second = True
                break
        tag_set = second_tag if use_second else first_tag
        tag_symbol, tag_name = _find_first_available_tag(doc, tag_set)

        if not tag_symbol:
            failed.append((elem, "No tag found from {} set".format("second" if use_second else "first")))
            continue

        # Skip if already tagged with the correct tag; otherwise delete wrong tags
        elem_id_val = elem.Id.IntegerValue
        existing_for_elem = tag_map.get(elem_id_val, [])
        if existing_for_elem:
            has_correct = False
            for existing_tag in existing_for_elem:
                existing_type = doc.GetElement(existing_tag.GetTypeId())
                if _tag_type_matches_target(existing_type, tag_name):
                    has_correct = True
                    break
            if has_correct:
                already_tagged.append(elem)
                continue

            for existing_tag in existing_for_elem:
                try:
                    doc.Delete(existing_tag.Id)
                except Exception:
                    pass

        # Get location point for tag placement - use element location directly
        tag_pt = None
        try:
            loc = elem.Location
            if hasattr(loc, 'Point'):
                tag_pt = loc.Point
        except Exception:
            pass

        if tag_pt is None:
            # Fallback to bounding box center
            view = uidoc.ActiveView
            bbox = elem.get_BoundingBox(view) if view else None
            if bbox:
                min_pt = bbox.Min
                max_pt = bbox.Max
                tag_pt = XYZ(
                    (min_pt.X + max_pt.X) / 2.0,
                    (min_pt.Y + max_pt.Y) / 2.0,
                    (min_pt.Z + max_pt.Z) / 2.0,
                )
            else:
                failed.append((elem, "Unable to determine tag location"))
                continue

        # Place tag using element directly with its location point
        try:
            tagger.place_tag(elem, tag_symbol, tag_pt)
            placed.append(elem)
        except Exception as e:
            failed.append((elem, "Tag placement error: {}".format(str(e))))

    t.Commit()
except Exception as e:
    # output.print_md("Tag placement error: {}".format(e))
    t.RollBack()
    raise

# Reporting
# ==================================================
# output.print_md(
#     "## Summary: placed {}, already tagged {}, failed {}".format(
#         len(placed),
#         len(already_tagged),
#         len(failed),
#     )
# )

# if failed:
#     output.print_md("\n### Failed Elements:")
#     for idx, (elem, reason) in enumerate(failed, 1):
#         output.print_md(
#             "### No: {:03} | ID: {} | Reason: {}".format(
#                 idx,
#                 output.linkify(elem.Id),
#                 reason
#             )
#         )
