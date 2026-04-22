# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from pyrevit import revit, script
from Autodesk.Revit.DB import (
    BuiltInCategory,
    ElementId,
    ElementType,
    FilteredElementCollector,
    FamilySymbol,
    IndependentTag,
    Transaction,
    XYZ,
)
from config.parameters_registry import *
from tagging.tag_config import (
    DEFAULT_PARAMETER_HIERARCHY,
    DEFAULT_TAG_SKIP_PARAMETERS,
    DEFAULT_TAG_SLOT_CANDIDATES,
    SLOT_GRD,
    SLOT_GRD_CFM,
    SLOT_LOUVER,
    WRITE_PARAMETER,
)

try:
    from tagging.revit_tagging import RevitTagging
except ImportError:
    from revit_tagging import RevitTagging

# Button display information
# =================================================
__title__ = "Tag Louvers"
__doc__ = """
Tags all air terminals in the current view
"""

# Helpers
# ==================================================

_AIR_TERMINAL_TAG_CATEGORIES = (
    BuiltInCategory.OST_DuctTerminalTags,
    BuiltInCategory.OST_MultiCategoryTags,
)


def _tag_candidate_parts(candidate):
    if not candidate:
        return "", "", ""
    if isinstance(candidate, tuple):
        family_name = str(candidate[0]).strip()
        type_name = str(candidate[1]).strip()
        pool = "{} {}".format(family_name, type_name).strip()
        return family_name, type_name, pool
    label = str(candidate).strip()
    return label, "", label


def _iter_tag_symbols(doc):
    seen_ids = set()
    for bic in _AIR_TERMINAL_TAG_CATEGORIES:
        symbols = (
            FilteredElementCollector(doc)
            .OfCategory(bic)
            .WhereElementIsElementType()
            .ToElements()
        )
        for sym in symbols:
            if not isinstance(sym, (FamilySymbol, ElementType)):
                continue
            try:
                symbol_id = sym.Id.IntegerValue
            except Exception:
                symbol_id = None
            if symbol_id is not None and symbol_id in seen_ids:
                continue
            if symbol_id is not None:
                seen_ids.add(symbol_id)
            yield sym


def _tag_symbol_family_type(symbol):
    if symbol is None:
        return "", ""

    try:
        fam_name, type_name, _ = tagger._tag_pool(symbol)
        return (fam_name or "").strip(), (type_name or "").strip()
    except Exception:
        pass

    fam = getattr(symbol, "Family", None)
    fam_name = fam.Name if fam else ""
    type_name = getattr(symbol, "Name", "") or ""
    return (fam_name or "").strip(), (type_name or "").strip()


def _find_tag_symbol(doc, target_candidate):
    """Return the first air terminal tag matching the configured candidate."""
    family_name, type_name, pool = _tag_candidate_parts(target_candidate)
    if not pool:
        return None

    if family_name and type_name:
        try:
            matched = tagger.get_label_exact(family_name, type_name)
            matched_family_name, matched_type_name = _tag_symbol_family_type(matched)
            if (
                matched_family_name.lower() == family_name.lower()
                and matched_type_name.lower() == type_name.lower()
            ):
                return matched
        except Exception:
            pass

        for sym in tagger.tag_syms:
            matched_family_name, matched_type_name = _tag_symbol_family_type(sym)
            if (
                matched_family_name.lower() == family_name.lower()
                and matched_type_name.lower() == type_name.lower()
            ):
                return sym

        for sym in _iter_tag_symbols(doc):
            matched_family_name, matched_type_name = _tag_symbol_family_type(sym)
            if (
                matched_family_name.lower() == family_name.lower()
                and matched_type_name.lower() == type_name.lower()
            ):
                return sym
        return None

    try:
        return tagger.get_label(pool)
    except Exception:
        pass

    family_needle = family_name.lower()
    type_needle = type_name.lower()
    pool_needle = pool.lower()
    exact_matches = []
    contains_matches = []
    for sym in _iter_tag_symbols(doc):
        fam = getattr(sym, "Family", None)
        fam_name = fam.Name if fam else ""
        symbol_type_name = getattr(sym, "Name", "") or ""
        fam_norm = fam_name.strip().lower()
        type_norm = symbol_type_name.strip().lower()
        label = (fam_name + " " + symbol_type_name).lower()
        if family_needle and type_needle:
            if family_needle == fam_norm and type_needle == type_norm:
                exact_matches.append(sym)
            elif family_needle in label and type_needle in label:
                contains_matches.append(sym)
        elif pool_needle == fam_norm or pool_needle == type_norm:
            exact_matches.append(sym)
        elif pool_needle in label:
            contains_matches.append(sym)

    if exact_matches:
        return exact_matches[0]
    if contains_matches:
        return contains_matches[0]
    return None


# Code
# ==================================================
uidoc = __revit__.ActiveUIDocument

doc = revit.doc
view = revit.active_view
output = script.get_output()

tagger = RevitTagging(doc, view)

# Tag sets in priority order
first_tag = list(DEFAULT_TAG_SLOT_CANDIDATES.get(SLOT_LOUVER, []))

second_tag = list(DEFAULT_TAG_SLOT_CANDIDATES.get(SLOT_LOUVER, []))

FLOW_PARAMETER_NAMES = (
    RVT_AIRFLOW,
    RVT_FLOW,
)

SUBJECT_PARAMETER_NAMES = (
    BBM_SUBJECT,
)

TARGET_SUBJECT_TEXT = "louver"

order_parameters = list(DEFAULT_PARAMETER_HIERARCHY)

value_parameters = {
    WRITE_PARAMETER,
}

skip_parameter_name = None
skip_parameter_values = set()
for _skip_param_name, _skip_values in DEFAULT_TAG_SKIP_PARAMETERS.items():
    skip_parameter_name = _skip_param_name
    skip_parameter_values = {str(v).lower().strip() for v in _skip_values}
    break


def _get_param_case_insensitive(element, param_name):
    target = param_name.lower().strip()
    for param in element.Parameters:
        try:
            if param.Definition and param.Definition.Name and param.Definition.Name.lower().strip() == target:
                return param
        except Exception:
            pass
    return None


def _parameter_has_positive_flow(param):
    if not param:
        return False

    try:
        storage_type = param.StorageType
        if storage_type == 1:
            return param.AsInteger() > 0
        if storage_type == 2:
            value = param.AsDouble()
            return value is not None and value > 1e-9

        value_text = param.AsString()
        if value_text is None:
            value_text = param.AsValueString()
        if not value_text:
            return False

        normalized = value_text.lower().strip().replace(',', '')
        token = normalized.split()[0]
        return float(token) > 0.0
    except Exception:
        return False


def _uses_no_flow_tag(element):
    for param_name in FLOW_PARAMETER_NAMES:
        param = _get_param_case_insensitive(element, param_name)
        if _parameter_has_positive_flow(param):
            return False
    return True


def _get_parameter_text(param):
    if not param:
        return ""
    try:
        val = param.AsString()
        if not val:
            val = param.AsValueString()
        return val.strip() if val else ""
    except Exception:
        return ""


def _get_subject_text(element):
    for param_name in SUBJECT_PARAMETER_NAMES:
        param = _get_param_case_insensitive(element, param_name)
        value = _get_parameter_text(param)
        if value:
            return value
    return ""


def _is_louver_subject(element):
    subject_text = _get_subject_text(element)
    return TARGET_SUBJECT_TEXT in subject_text.lower()


def _find_first_available_tag(doc, tag_names):
    """Try to find the first available tag from a list of configured candidates."""
    for tag_candidate in tag_names:
        tag_sym = _find_tag_symbol(doc, tag_candidate)
        if tag_sym:
            return tag_sym, tag_candidate
    return None, None


def _format_tag_candidate(candidate):
    if isinstance(candidate, tuple):
        return "{} :: {}".format(candidate[0], candidate[1])
    return str(candidate)


def _format_tag_symbol(symbol):
    if symbol is None:
        return "none"
    fam_name, type_name = _tag_symbol_family_type(symbol)
    return "{} :: {}".format(fam_name or '<no family>', type_name or '<no type>')


def _tag_symbol_matches_candidate(symbol, candidate):
    if symbol is None or candidate is None:
        return False
    family_name, type_name, pool = _tag_candidate_parts(candidate)
    sym_family_name, sym_type_name = _tag_symbol_family_type(symbol)
    sym_family_name = sym_family_name.lower()
    sym_type_name = sym_type_name.lower()
    if family_name and type_name:
        return (
            sym_family_name == family_name.lower()
            and sym_type_name == type_name.lower()
        )
    pool_lower = pool.lower()
    label = "{} {}".format(sym_family_name, sym_type_name).strip()
    return pool_lower == sym_family_name or pool_lower == sym_type_name or pool_lower in label


def _eid_int(eid):
    try:
        return eid.IntegerValue
    except Exception:
        try:
            return int(eid)
        except Exception:
            return None


def _collect_tagged_local_ids(tag):
    ids = []

    try:
        for tid in tag.GetTaggedLocalElementIds() or []:
            if tid and tid != ElementId.InvalidElementId:
                ids.append(tid)
    except Exception:
        pass

    try:
        tid = tag.TaggedLocalElementId
        if tid and tid != ElementId.InvalidElementId:
            ids.append(tid)
    except Exception:
        pass

    def _append_from_link_eid(link_eid):
        if not link_eid:
            return
        for attr in ("HostElementId", "LinkedElementId", "ElementId"):
            try:
                candidate = getattr(link_eid, attr)
                if candidate and candidate != ElementId.InvalidElementId:
                    ids.append(candidate)
            except Exception:
                pass

    try:
        for leid in tag.GetTaggedElementIds() or []:
            _append_from_link_eid(leid)
    except Exception:
        pass

    try:
        _append_from_link_eid(tag.TaggedElementId)
    except Exception:
        pass

    unique_ids = []
    seen = set()
    for eid in ids:
        eid_int = _eid_int(eid)
        if eid_int is None or eid_int in seen:
            continue
        seen.add(eid_int)
        unique_ids.append(eid)
    return unique_ids


def _tag_targets_element(tag, element_id_int):
    for tagged_id in _collect_tagged_local_ids(tag):
        if _eid_int(tagged_id) == element_id_int:
            return True
    return False


def _is_grd_tag_type(tag_type):
    if tag_type is None:
        return False
    for candidate in first_tag + second_tag:
        if _tag_symbol_matches_candidate(tag_type, candidate):
            return True
    return False


def _get_value_from_ordered_params(element, doc, param_order):
    """Get value from element following the ordered parameter hierarchy.
    Returns the last non-empty value found, or empty string if none found.
    Does case-insensitive parameter lookup.
    """
    elem_type = doc.GetElement(element.GetTypeId())
    result = ""

    for param_name in param_order:
        param_name_lower = param_name.lower().strip()

        # Try instance parameters with case-insensitive lookup
        for param in element.Parameters:
            if param.Definition.Name.lower().strip() == param_name_lower:
                try:
                    val = param.AsString()
                    if not val:
                        val = param.AsValueString()
                    if val and val.strip():
                        result = val.strip()
                        break
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
                            result = val.strip()
                            break
                    except Exception:
                        pass

    return result


def _is_element_visible_in_view(element, active_view):
    if element is None or active_view is None:
        return False
    try:
        bbox = element.get_BoundingBox(active_view)
        if bbox is not None:
            return True
    except Exception:
        pass
    return False


def _collect_air_terminals(doc, active_view):
    view_elements = list(
        FilteredElementCollector(doc, active_view.Id)
        .OfCategory(BuiltInCategory.OST_DuctTerminal)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    if view_elements:
        return view_elements, False, None

    all_elements = list(
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_DuctTerminal)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    visible_elements = [
        elem for elem in all_elements
        if _is_element_visible_in_view(elem, active_view)
    ]
    if visible_elements:
        return visible_elements, True, len(all_elements)
    return [], True, len(all_elements)


air_terminals, used_fallback_collection, total_air_terminals = _collect_air_terminals(doc, view)
subject_filtered_air_terminals = [
    elem for elem in air_terminals
    if _is_louver_subject(elem)
]

if not air_terminals:
    if used_fallback_collection:
        output.print_md(
            "## No GRDs found in this view. Active view: {} ({}) | Air terminals in model: {}".format(
                getattr(view, 'Name', '<unknown>'),
                view.ViewType,
                total_air_terminals or 0,
            )
        )
    else:
        output.print_md("## No GRDs found in this view.")
    script.exit()

if not subject_filtered_air_terminals:
    output.print_md(
        "## No GRDs in this view have {} containing '{}'.".format(
            SUBJECT_PARAMETER_NAMES[0],
            TARGET_SUBJECT_TEXT,
        )
    )
    script.exit()

air_terminals = subject_filtered_air_terminals

if used_fallback_collection:
    output.print_md(
        "## View collector returned 0 GRDs; using bounding-box fallback from {} model air terminals and found {} louver candidates in view.".format(
            total_air_terminals or 0,
            len(air_terminals),
        ))


first_tag_symbol, first_tag_name = _find_first_available_tag(doc, first_tag)
second_tag_symbol, second_tag_name = _find_first_available_tag(doc, second_tag)

if not first_tag_symbol and not second_tag_symbol:
    output.print_md(
        "## No GRD tag types resolved. With-flow candidates: {} | No-flow candidates: {}".format(
            ", ".join(_format_tag_candidate(candidate) for candidate in first_tag) or "none",
            ", ".join(_format_tag_candidate(candidate) for candidate in second_tag) or "none",
        )
    )
    script.exit()
placed = []
failed = []
already_tagged = []
skipped = []
value_changes = []
tags_deleted = []
tags_added = []

# Check how many tags already exist in the view
existing_tags = list(
    FilteredElementCollector(doc, view.Id)
    .OfClass(IndependentTag)
    .ToElements()
)

# Build maps of element IDs keyed by tracked tag type id.
existing_tag_maps = {}
tracked_tag_type_ids = set()
if first_tag_symbol is not None:
    tracked_tag_type_ids.add(first_tag_symbol.Id.IntegerValue)
if second_tag_symbol is not None:
    tracked_tag_type_ids.add(second_tag_symbol.Id.IntegerValue)

for tag in existing_tags:
    try:
        tag_type_id = tag.GetTypeId()
        tag_type_id_val = _eid_int(tag_type_id)
        if tag_type_id is None or tag_type_id_val not in tracked_tag_type_ids:
            continue

        for tid in _collect_tagged_local_ids(tag):
            tid_val = _eid_int(tid)
            if tid_val is None or tid_val == -1:
                continue
            existing_tag_maps.setdefault(tag_type_id_val, {}).setdefault(tid_val, []).append(tag)
    except BaseException:
        pass

t = Transaction(doc, "Tag Air Terminals")
t.Start()
try:
    skip_values_normalized = set(skip_parameter_values)

    # Update/tag air terminals
    for elem in air_terminals:
        # Check configured skip parameter/value pair
        if skip_parameter_name:
            skip_param = _get_param_case_insensitive(elem, skip_parameter_name)
            skip_value = _get_parameter_text(skip_param)
            skip_value_normalized = skip_value.lower().strip()

            if skip_value_normalized in skip_values_normalized:
                elem_id_val = elem.Id.IntegerValue

                # Remove any existing GRD tags for skipped elements by scanning tags in view.
                for existing_tag in list(existing_tags):
                    try:
                        if not _tag_targets_element(existing_tag, elem_id_val):
                            continue
                        tag_type = doc.GetElement(existing_tag.GetTypeId())
                        if not _is_grd_tag_type(tag_type):
                            continue

                        doc.Delete(existing_tag.Id)
                        tags_deleted.append((
                            elem,
                            existing_tag.Id,
                            _format_tag_symbol(tag_type),
                            "skip-cleanup",
                        ))
                    except Exception:
                        pass

                # Clear any cached map entries for this element.
                for tracked_type_id in tracked_tag_type_ids:
                    existing_tag_maps.setdefault(tracked_type_id, {})[elem_id_val] = []

                # Clear write parameter for skipped elements.
                for param_name in value_parameters:
                    try:
                        param = _get_param_case_insensitive(elem, param_name)
                        if not param:
                            continue
                        current_value = _get_parameter_text(param)
                        if not current_value:
                            continue
                        param.Set("")
                        value_changes.append((
                            elem,
                            param_name,
                            current_value,
                            "",
                        ))
                    except Exception:
                        pass

                skipped.append((
                    elem,
                    "Parameter '{}' has skip value '{}'".format(skip_parameter_name, skip_value)
                ))
                continue

        # Get value based on ordered parameter hierarchy
        value_to_write = _get_value_from_ordered_params(elem, doc, order_parameters)

        # Write only when target value parameter is currently empty.
        for param_name in value_parameters:
            try:
                param = _get_param_case_insensitive(elem, param_name)
                if not param:
                    continue

                current_value = _get_parameter_text(param)
                if not value_to_write:
                    continue

                if current_value == value_to_write:
                    continue

                param.Set(value_to_write)
                value_changes.append((
                    elem,
                    param_name,
                    current_value,
                    value_to_write,
                ))
            except Exception:
                pass

        # Use the no-flow tag only when neither airflow parameter has a positive value.
        use_second = _uses_no_flow_tag(elem)
        tag_symbol = second_tag_symbol if use_second else first_tag_symbol
        tag_name = second_tag_name if use_second else first_tag_name

        if not tag_symbol:
            failed.append((
                elem,
                "No tag found from {} set ({})".format(
                    "no-flow" if use_second else "with-flow",
                    ", ".join(
                        _format_tag_candidate(candidate)
                        for candidate in (second_tag if use_second else first_tag)
                    ) or "none",
                )
            ))
            continue

        # Skip if already tagged with the correct tag; otherwise delete wrong tags
        elem_id_val = elem.Id.IntegerValue
        chosen_type_id_val = _eid_int(tag_symbol.Id)
        existing_for_elem = existing_tag_maps.get(chosen_type_id_val, {}).get(elem_id_val, [])
        if existing_for_elem:
            already_tagged.append(elem)
            continue

        opposite_type_id_val = None
        if first_tag_symbol is not None and second_tag_symbol is not None:
            if chosen_type_id_val == _eid_int(first_tag_symbol.Id):
                opposite_type_id_val = _eid_int(second_tag_symbol.Id)
            elif chosen_type_id_val == _eid_int(second_tag_symbol.Id):
                opposite_type_id_val = _eid_int(first_tag_symbol.Id)

        if opposite_type_id_val is not None:
            wrong_tags = list(existing_tag_maps.get(opposite_type_id_val, {}).get(elem_id_val, []))
            for existing_tag in wrong_tags:
                try:
                    deleted_type = doc.GetElement(existing_tag.GetTypeId())
                    doc.Delete(existing_tag.Id)
                    tags_deleted.append((
                        elem,
                        existing_tag.Id,
                        _format_tag_symbol(deleted_type),
                        "replaced-by-correct-type",
                    ))
                except Exception:
                    pass
            existing_tag_maps.setdefault(opposite_type_id_val, {})[elem_id_val] = []

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
            new_tag = tagger.place_tag(elem, tag_symbol, tag_pt)
            if new_tag is None:
                reason = tagger.last_place_tag_failure or "Tag placement returned no tag"
                failed.append((elem, "Tag placement error [{}]: {}".format(tag_name, reason)))
                continue

            placed_type = doc.GetElement(new_tag.GetTypeId())
            if not _tag_symbol_matches_candidate(placed_type, tag_name):
                try:
                    new_tag.ChangeTypeId(tag_symbol.Id)
                    placed_type = doc.GetElement(new_tag.GetTypeId())
                except Exception:
                    pass

            if not _tag_symbol_matches_candidate(placed_type, tag_name):
                actual_type = _format_tag_symbol(placed_type)
                expected_type = _format_tag_candidate(tag_name)
                try:
                    bad_tag_id = new_tag.Id
                    doc.Delete(new_tag.Id)
                    tags_deleted.append((
                        elem,
                        bad_tag_id,
                        actual_type,
                        "placement-mismatch-cleanup",
                    ))
                except Exception:
                    pass
                failed.append((
                    elem,
                    "Placed tag type mismatch. Expected [{}], got [{}]".format(
                        expected_type,
                        actual_type,
                    )
                ))
                continue

            placed.append(elem)
            existing_tag_maps.setdefault(chosen_type_id_val, {}).setdefault(elem_id_val, []).append(new_tag)
            tags_added.append((
                elem,
                new_tag.Id,
                _format_tag_symbol(placed_type),
            ))
        except Exception as e:
            failed.append((elem, "Tag placement error [{}]: {}".format(tag_name, str(e))))

    t.Commit()
except Exception as e:
    # output.print_md("Tag placement error: {}".format(e))
    t.RollBack()
    raise

# Reporting
# ==================================================
output.print_md(
    "## Summary: placed {}, already tagged {}, skipped {}, failed {}".format(
        len(placed),
        len(already_tagged),
        len(skipped),
        len(failed),
    )
)

output.print_md(
    "## Tag Candidates: with airflow {}, without airflow {}".format(
        ", ".join(_format_tag_candidate(candidate) for candidate in first_tag) or "none",
        ", ".join(_format_tag_candidate(candidate) for candidate in second_tag) or "none",
    )
)

output.print_md(
    "## Resolved Tags: with airflow {}, without airflow {}".format(
        _format_tag_symbol(first_tag_symbol),
        _format_tag_symbol(second_tag_symbol),
    )
)

output.print_md("## Values Changed: {}".format(len(value_changes)))
if value_changes:
    for idx, (elem, param_name, old_value, new_value) in enumerate(value_changes, 1):
        output.print_md(
            "- {:03} | ID: {} | Param: {} | Old: '{}' | New: '{}'".format(
                idx,
                output.linkify(elem.Id),
                param_name,
                old_value,
                new_value,
            )
        )

output.print_md("## Tags Deleted: {}".format(len(tags_deleted)))
if tags_deleted:
    for idx, (elem, tag_id, tag_type_name, reason) in enumerate(tags_deleted, 1):
        output.print_md(
            "- {:03} | Elem ID: {} | Tag ID: {} | Type: {} | Reason: {}".format(
                idx,
                output.linkify(elem.Id),
                output.linkify(tag_id),
                tag_type_name,
                reason,
            )
        )

output.print_md("## Tags Added: {}".format(len(tags_added)))
if tags_added:
    for idx, (elem, tag_id, tag_type_name) in enumerate(tags_added, 1):
        output.print_md(
            "- {:03} | Elem ID: {} | Tag ID: {} | Type: {}".format(
                idx,
                output.linkify(elem.Id),
                output.linkify(tag_id),
                tag_type_name,
            )
        )

if skipped:
    output.print_md("\n### Skipped Elements:")
    for idx, (elem, reason) in enumerate(skipped, 1):
        output.print_md(
            "- {:03} | ID: {} | Reason: {}".format(
                idx,
                output.linkify(elem.Id),
                reason
            )
        )

if failed:
    output.print_md("\n### Failed Elements:")
    for idx, (elem, reason) in enumerate(failed, 1):
        output.print_md(
            "- {:03} | ID: {} | Reason: {}".format(
                idx,
                output.linkify(elem.Id),
                reason
            )
        )
