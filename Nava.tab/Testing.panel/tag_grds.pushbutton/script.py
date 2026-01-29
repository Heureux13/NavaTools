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
    BuiltInParameter,
    FilteredElementCollector,
    FamilySymbol,
    IndependentTag,
    Transaction,
)
from revit_tagging import RevitTagging

# Button display information
# =================================================
__title__ = "Tag Air Terminals"
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
    for sym in symbols:
        fam = getattr(sym, "Family", None)
        fam_name = fam.Name if fam else ""
        type_name = getattr(sym, "Name", "") or ""
        label = (fam_name + " " + type_name).lower()
        if needle in label:
            return sym
    return None


# Code
# ==================================================
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

doc = revit.doc
view = revit.active_view
output = script.get_output()

tagger = RevitTagging(doc, view)

target_tag_name = "-UMI_GRD_JN"
tag_symbol = _find_tag_symbol(doc, target_tag_name)

if not tag_symbol:
    output.print_md("## Tag '{}' not found in Air Terminal Tags.".format(target_tag_name))
    script.exit()

air_terminals = (
    FilteredElementCollector(doc, view.Id)
    .OfCategory(BuiltInCategory.OST_DuctTerminal)
    .WhereElementIsNotElementType()
    .ToElements()
)

if not air_terminals:
    output.print_md("## No air terminals found in this view.")
    script.exit()

placed = []
failed = []
already_tagged = []

fam_name = (tag_symbol.Family.Name if tag_symbol and tag_symbol.Family else "").strip()

# Check how many tags already exist in the view
existing_tags = list(
    FilteredElementCollector(doc, view.Id)
    .OfClass(IndependentTag)
    .ToElements()
)

# Build a set of element IDs that are already tagged with our tag family
already_tagged_ids = set()
fam_name_lower = fam_name.strip().lower()

for tag in existing_tags:
    try:
        # Check if this tag is using our family
        tag_type_id = tag.GetTypeId()
        tag_type = doc.GetElement(tag_type_id)
        if tag_type and hasattr(tag_type, 'Family'):
            tag_fam_name = (tag_type.Family.Name or "").strip().lower()
            if tag_fam_name == fam_name_lower:
                # This tag uses our family, get what it's tagging
                tagged_ids = tag.GetTaggedLocalElementIds()
                for tid in tagged_ids:
                    already_tagged_ids.add(tid)
    except BaseException:
        pass

t = Transaction(doc, "Tag Air Terminals")
t.Start()
try:
    # Update _grd_label parameter for all air terminals before tagging
    for elem in air_terminals:
        try:
            # Get instance parameters
            grd_label_instance_param = elem.LookupParameter("_grd_label_instance")
            mark_param = elem.LookupParameter("Mark")
            grd_label_param = elem.LookupParameter("_grd_label")

            # Get type parameters
            elem_type = doc.GetElement(elem.GetTypeId())
            grd_label_type_param = elem_type.LookupParameter("_grd_label_type") if elem_type else None
            type_mark_param = elem_type.LookupParameter("Type Mark") if elem_type else None

            # Determine value to write based on priority
            value_to_write = ""

            if grd_label_instance_param:
                val = grd_label_instance_param.AsString()
                if val and val.strip():
                    value_to_write = val

            if not value_to_write and mark_param:
                val = mark_param.AsString()
                if val and val.strip():
                    value_to_write = val

            if not value_to_write and grd_label_type_param:
                val = grd_label_type_param.AsString()
                if val and val.strip():
                    value_to_write = val

            if not value_to_write and type_mark_param:
                val = type_mark_param.AsString()
                if val and val.strip():
                    value_to_write = val

            # Always write to _grd_label (even if empty)
            if grd_label_param:
                grd_label_param.Set(value_to_write)
        except Exception:
            pass

    # Now tag all air terminals
    for elem in air_terminals:
        # Check if already tagged with our tag family
        if elem.Id in already_tagged_ids:
            already_tagged.append(elem)
            continue

        # Try a face reference first, then fall back to element center
        try:
            face_ref, face_pt = tagger.get_face_facing_view(elem)
        except Exception:
            face_ref, face_pt = (None, None)

        placed_one = False

        # Attempt face placement
        if face_ref is not None and face_pt is not None:
            try:
                tagger.place_tag(face_ref, tag_symbol, face_pt)
                placed.append(elem)
                placed_one = True
            except Exception:
                pass

        # Fallback: place at element center
        if not placed_one:
            try:
                bbox = elem.get_BoundingBox(view)
                if bbox:
                    center = (bbox.Min + bbox.Max) / 2.0
                    tagger.place_tag(elem, tag_symbol, center)
                    placed.append(elem)
                    placed_one = True
            except Exception:
                pass

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
    "## Summary: placed {}, already tagged {}, failed {}".format(
        len(placed),
        len(already_tagged),
        len(failed),
    )
)
