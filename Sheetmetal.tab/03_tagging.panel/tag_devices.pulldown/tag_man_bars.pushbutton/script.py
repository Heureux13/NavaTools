# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from pyrevit import DB, revit, script
from Autodesk.Revit.DB import (
    BuiltInCategory,
    FilteredElementCollector,
    Transaction,
    XYZ,
)
from tagging.revit_tagging import RevitTagging
from tagging.revit_tagging_fittings import Fittings

# Button display information
# =================================================
__title__ = "Tag Man Bars"
__doc__ = """
Tags only ManBars fabrication ductwork in the current view with MARK label tags.
"""


TARGET_FAMILY_NAMES = {"manbars"}


def _normalized(value):
    return (str(value or "").strip().lower())


def _get_param_value(elem, param_name):
    target = _normalized(param_name)
    if not target or elem is None:
        return None

    try:
        for p in elem.Parameters:
            try:
                dname = p.Definition.Name if p and p.Definition else None
                if _normalized(dname) != target:
                    continue
                val = p.AsString() or p.AsValueString()
                if val:
                    return str(val).strip()
            except Exception:
                pass
    except Exception:
        pass

    return None


def _get_element_family_name(elem, doc):
    # Try common family surfaces first.
    try:
        fam = getattr(elem, "Family", None)
        if fam and getattr(fam, "Name", None):
            return str(fam.Name).strip()
    except Exception:
        pass

    try:
        sym = getattr(elem, "Symbol", None)
        sym_fam = getattr(sym, "Family", None) if sym else None
        if sym_fam and getattr(sym_fam, "Name", None):
            return str(sym_fam.Name).strip()
    except Exception:
        pass

    # Fabrication parts often expose family info through parameters.
    for pname in ("Family", "Family Name", "Fabrication Fitting Description"):
        val = _get_param_value(elem, pname)
        if val:
            return val

    # Last fallback: inspect element type.
    try:
        type_elem = doc.GetElement(elem.GetTypeId())
        if type_elem is not None:
            for pname in ("Family", "Family Name", "Fabrication Fitting Description"):
                val = _get_param_value(type_elem, pname)
                if val:
                    return val
    except Exception:
        pass

    return ""


# Code
# ==================================================
uidoc = __revit__.ActiveUIDocument

doc = revit.doc
view = revit.active_view
output = script.get_output()

tagger = RevitTagging(doc, view)
fittings = Fittings(doc=doc, view=view, tagger=tagger)


equipment_elements = (
    FilteredElementCollector(doc, view.Id)
    .OfCategory(BuiltInCategory.OST_FabricationDuctwork)
    .WhereElementIsNotElementType()
    .ToElements()
)

equipment_elements = [
    e for e in equipment_elements
    if _normalized(_get_element_family_name(e, doc)) in TARGET_FAMILY_NAMES
]

if not equipment_elements:
    output.print_md("## No ManBars fabrication ductwork found in this view.")
    script.exit()

selected_tag_symbol, selected_tag_name = fittings._resolve_slot(fittings.SLOT_MARK)
if not selected_tag_symbol:
    output.print_md(
        "## No MARK tag found from tag_map slot: {}".format(fittings.SLOT_MARK)
    )
    script.exit()

placed = []
failed = []
already_tagged = []

selected_tag_family_name = (
    selected_tag_symbol.Family.Name.lower()
    if selected_tag_symbol and selected_tag_symbol.Family
    else ""
)
existing_tag_family_map = tagger.build_existing_tag_family_map(equipment_elements)

t = Transaction(doc, "Tag ManBars Fabrication Ductwork MARK")
t.Start()
try:
    for elem in equipment_elements:
        try:
            if not elem.Category or elem.Category.Id.IntegerValue != int(BuiltInCategory.OST_FabricationDuctwork):
                failed.append((elem, "Skipped non-fabrication ductwork category"))
                continue
        except Exception:
            failed.append((elem, "Unable to validate category"))
            continue

        fittings.update_write_parameter_from_hierarchy(elem)

        tag_symbol = selected_tag_symbol

        # Skip only if element already has this MARK tag in this view
        elem_id_val = elem.Id.IntegerValue
        existing_fams = existing_tag_family_map.get(elem_id_val, set())
        if selected_tag_family_name and selected_tag_family_name in existing_fams:
            already_tagged.append(elem)
            continue

        # First try rotation-aware placement so tags follow the duct direction.
        # If that is not possible for the element, fallback to face/point placement.
        tag_pt = None
        try:
            placed_tag = False
            rotated_tag = tagger.place_tag_at_center_with_rotation(
                elem,
                tag_label=tag_symbol,
                position="center",
            )
            if rotated_tag is not None:
                placed_tag = True

            if (not placed_tag) and isinstance(elem, DB.FabricationPart):
                face_ref, face_pt = tagger.get_face_facing_view(
                    elem, prefer_point=None)
                if face_ref is not None and face_pt is not None:
                    tagger.place_tag(face_ref, tag_symbol, face_pt)
                    placed_tag = True

            if not placed_tag:
                loc = getattr(elem, "Location", None)
                if hasattr(loc, "Point") and loc.Point is not None:
                    tag_pt = loc.Point
                elif hasattr(loc, "Curve") and loc.Curve is not None:
                    tag_pt = loc.Curve.Evaluate(0.5, True)
                else:
                    active_view = uidoc.ActiveView
                    bbox = elem.get_BoundingBox(active_view) if active_view else None
                    if bbox:
                        min_pt = bbox.Min
                        max_pt = bbox.Max
                        tag_pt = XYZ(
                            (min_pt.X + max_pt.X) / 2.0,
                            (min_pt.Y + max_pt.Y) / 2.0,
                            (min_pt.Z + max_pt.Z) / 2.0,
                        )

                if tag_pt is None:
                    failed.append((elem, "Unable to determine tag location"))
                    continue

                tagger.place_tag(elem, tag_symbol, tag_pt)

            placed.append(elem)
            if selected_tag_family_name:
                existing_tag_family_map.setdefault(elem_id_val, set()).add(
                    selected_tag_family_name
                )
        except Exception as e:
            failed.append((elem, "Tag placement error: {}".format(str(e))))

    t.Commit()
except Exception as e:
    # output.print_md("Tag placement error: {}".format(e))
    t.RollBack()
    raise

output.print_md(
    "## Summary: placed {}, already tagged {}, failed {}".format(
        len(placed),
        len(already_tagged),
        len(failed),
    )
)

if failed:
    output.print_md("\n### Failed Elements:")
    for idx, (elem, reason) in enumerate(failed, 1):
        output.print_md(
            "- {:03} | ID: {} | Reason: {}".format(
                idx,
                output.linkify(elem.Id),
                reason,
            )
        )
