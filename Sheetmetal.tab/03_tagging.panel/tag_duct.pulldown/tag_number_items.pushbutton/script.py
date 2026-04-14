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
    FilteredElementCollector,
    FamilySymbol,
    IndependentTag,
    Transaction,
    StorageType,
)
from tagging.revit_tagging import RevitTagging
from config.parameters_registry import PYT_NUMBER_FABRICATION, RVT_FAMILY, RVT_ITEM_NUMBER
from tagging.tag_config import (
    DEFAULT_NUMBER_SKIP_PARAMETERS,
    DEFAULT_TAG_SLOT_CANDIDATES,
    SLOT_NUMBER_FABRICATION,
)

# Button display information
# =================================================
__title__ = "Tag Item Number All"
__doc__ = """
Tags Item Number
"""

# Helpers
# ==================================================


def _find_tag_symbol(doc, target_name):
    """Return the first fabrication duct tag whose name contains target_name."""
    if not target_name:
        return None
    needle = target_name.strip().lower()
    symbols = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_FabricationDuctworkTags)
        .OfClass(FamilySymbol)
        .ToElements()
    )
    for sym in symbols:
        fam = getattr(sym, "Family", None)
        fam_name = fam.Name if fam else ""
        type_name = getattr(sym, "Name", "") or ""
        label = (fam_name + " " + type_name).lower()
        if needle in label:
            return sym
    return None


def _eid_int(eid):
    """Normalize ElementId and LinkElementId-like objects to an integer value."""
    if eid is None:
        return None

    host_id = getattr(eid, "HostElementId", None)
    if host_id is not None:
        eid = host_id

    for attr in ("Value", "IntegerValue"):
        try:
            value = getattr(eid, attr)
            if value is not None:
                return int(value)
        except Exception:
            pass

    try:
        return int(eid)
    except Exception:
        return None


def _collect_tagged_local_ids(tag):
    """Return tagged element ids for an IndependentTag across Revit versions."""
    ids = []

    try:
        for tid in tag.GetTaggedLocalElementIds() or []:
            ids.append(tid)
    except Exception:
        pass

    try:
        tid = tag.TaggedLocalElementId
        if tid is not None:
            ids.append(tid)
    except Exception:
        pass

    try:
        for tid in tag.GetTaggedElementIds() or []:
            ids.append(tid)
    except Exception:
        pass

    try:
        tid = tag.TaggedElementId
        if tid is not None:
            ids.append(tid)
    except Exception:
        pass

    unique_ids = []
    seen = set()
    for tid in ids:
        tid_int = _eid_int(tid)
        if tid_int is None or tid_int in seen:
            continue
        seen.add(tid_int)
        unique_ids.append(tid)
    return unique_ids


def _get_tag_family_name(doc, tag):
    """Return the lowercase family name for a tag instance."""
    try:
        tag_type = doc.GetElement(tag.GetTypeId())
        if tag_type and hasattr(tag_type, "Family") and tag_type.Family:
            return (tag_type.Family.Name or "").strip().lower()
    except Exception:
        pass
    return ""


def _get_parameter_value(param):
    """Return a parameter value as a normalized string."""
    if not param:
        return ""

    value = param.AsString()
    if value is None:
        value = param.AsValueString()
    if value is None:
        if param.StorageType == StorageType.Integer:
            value = str(param.AsInteger())
        elif param.StorageType == StorageType.Double:
            value = str(param.AsDouble())

    return (value or "").strip()


def _get_first_matching_parameter(elem, parameter_names):
    """Return the first existing parameter from parameter_names in priority order."""
    for parameter_name in parameter_names:
        param = elem.LookupParameter(parameter_name)
        if param:
            return param
    return None


# Code
# ==================================================
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

doc = revit.doc
view = revit.active_view
output = script.get_output()

families_to_tag = {
    "straight",
    "canvas",
    "boot tap",
    "transition",
    "elbow - 90 degree",
    "elbow",
    "drop cheeck",
    "ogee",
    "offset",
    "square to Ø",
    "end cap",
    "tdf end cap",
    'reducer',
    'conical tee',
}

tagger = RevitTagging(doc, view)

number_parameter_names = [
    PYT_NUMBER_FABRICATION,
    RVT_ITEM_NUMBER,
]

tag_names = [
    family_name
    for family_name, _ in DEFAULT_TAG_SLOT_CANDIDATES.get(
        SLOT_NUMBER_FABRICATION, []
    )
]

values_to_skip = {
    "0",
    "skip",
}

for skip_values in DEFAULT_NUMBER_SKIP_PARAMETERS.values():
    for skip_value in skip_values:
        values_to_skip.add(str(skip_value))

families_to_tag_norm = {v.strip().lower() for v in families_to_tag if v}
values_to_skip_norm = {v.strip().lower() for v in values_to_skip if v}

tag_symbol = None
target_tag_name = None
candidate_tag_family_names = set()
for candidate in tag_names:
    matched_symbol = _find_tag_symbol(doc, candidate)
    if matched_symbol and matched_symbol.Family:
        candidate_tag_family_names.add(
            (matched_symbol.Family.Name or "").strip().lower()
        )
    if matched_symbol and not tag_symbol:
        tag_symbol = matched_symbol
        target_tag_name = candidate

if not tag_symbol:
    output.print_md(
        "## None of these tags were found in Fabrication Duct Tags: {}".format(
            ", ".join(tag_names)
        )
    )
    script.exit()

fab_ducts = (
    FilteredElementCollector(doc, view.Id)
    .OfCategory(BuiltInCategory.OST_FabricationDuctwork)
    .WhereElementIsNotElementType()
    .ToElements()
)


if not fab_ducts:
    output.print_md("## No fabrication ducts found in this view.")
    script.exit()

placed = []
failed = []
already_tagged = []
removed = []

fam_name = (
    tag_symbol.Family.Name if tag_symbol and tag_symbol.Family else "").strip()

# Check how many tags already exist in the view
existing_tags = list(
    FilteredElementCollector(doc, view.Id)
    .OfClass(IndependentTag)
    .ToElements()
)

# Build a map of element ID -> list of tag IDs using our tag family
already_tagged_ids = set()
elem_to_tags = {}
fam_name_lower = fam_name.strip().lower()

for tag in existing_tags:
    try:
        tag_fam_name = _get_tag_family_name(doc, tag)
        if tag_fam_name not in candidate_tag_family_names:
            continue

        tagged_ids = _collect_tagged_local_ids(tag)
        for tid in tagged_ids:
            tid_int = _eid_int(tid)
            if tid_int is None:
                continue

            if tag_fam_name == fam_name_lower:
                already_tagged_ids.add(tid_int)
            elem_to_tags.setdefault(tid_int, []).append(tag)
    except BaseException:
        pass

t = Transaction(doc, "Tag Item Number")
t.Start()
try:
    # Tag fabrication ducts (exclude certain families)
    for elem in fab_ducts:
        # Only tag elements whose family is in families_to_tag
        try:
            fam_param = elem.LookupParameter(RVT_FAMILY)
            if not fam_param:
                continue
            fam_value = _get_parameter_value(fam_param).lower()
            if not any(f in fam_value for f in families_to_tag_norm):
                continue
        except Exception:
            continue

        # Skip elements with empty or disallowed item number values
        try:
            item_param = _get_first_matching_parameter(elem, number_parameter_names)
            if not item_param:
                continue
            item_value = _get_parameter_value(item_param).lower()
            if not item_value or item_value in values_to_skip_norm:
                # Remove any existing tags for this element
                for existing_tag in elem_to_tags.get(elem.Id.IntegerValue, []):
                    try:
                        doc.Delete(existing_tag.Id)
                        removed.append(elem)
                    except Exception:
                        pass
                continue
        except Exception:
            pass

        existing_for_elem = list(elem_to_tags.get(elem.Id.IntegerValue, []))
        kept_existing_tag = None
        for existing_tag in existing_for_elem:
            tag_family_name = _get_tag_family_name(doc, existing_tag)
            if tag_family_name == fam_name_lower and kept_existing_tag is None:
                kept_existing_tag = existing_tag
                continue

            try:
                doc.Delete(existing_tag.Id)
                removed.append(elem)
            except Exception:
                pass

        # Check if already tagged with our tag family after pruning duplicates/conflicts
        if kept_existing_tag is not None or elem.Id.IntegerValue in already_tagged_ids:
            already_tagged.append(elem)
            continue

        placed_one = False

        # Place tag at element center with rotation
        if not placed_one:
            try:
                bbox = elem.get_BoundingBox(view)
                if bbox:
                    center = (bbox.Min + bbox.Max) / 2.0
                    tag = tagger.place_tag(elem, tag_symbol, center)

                    # Rotate tag to match duct direction from curve
                    try:
                        loc = elem.Location
                        if loc and hasattr(loc, "Curve") and loc.Curve:
                            import math
                            from Autodesk.Revit.DB import Line, ElementTransformUtils, XYZ

                            # Get the curve direction
                            curve = loc.Curve
                            dir_vec = (curve.GetEndPoint(1) -
                                       curve.GetEndPoint(0)).Normalize()

                            # Calculate angle in XY plane (plan view)
                            angle_rad = math.atan2(dir_vec.Y, dir_vec.X)

                            # Rotate tag to match duct direction
                            axis = Line.CreateBound(
                                tag.TagHeadPosition,
                                XYZ(tag.TagHeadPosition.X, tag.TagHeadPosition.Y,
                                    tag.TagHeadPosition.Z + 1)
                            )
                            ElementTransformUtils.RotateElement(
                                doc, tag.Id, axis, angle_rad)
                    except Exception:
                        pass

                    placed.append(elem)
                    placed_one = True
            except Exception as e:
                failed.append((elem, "Placement failed: {}".format(str(e))))

        if not placed_one:
            failed.append((elem, "No valid reference or center placement"))

    t.Commit()
except Exception as e:
    output.print_md("Tag placement error: {}".format(e))
    t.RollBack()
    raise

# Reporting
# ==================================================
output.print_md(
    "## Summary: placed {}, already tagged {}, removed {}, failed {}".format(
        len(placed),
        len(already_tagged),
        len(removed),
        len(failed),
    )
)
