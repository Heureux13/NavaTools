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
from revit_tagging import RevitTagging
from revit_duct import RevitDuct

# Button display information
# =================================================
__title__ = "Tag Item Number"
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


# Code
# ==================================================
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

doc = revit.doc
view = revit.active_view
output = script.get_output()

tagger = RevitTagging(doc, view)

tag_names = [
    "-FabDuct_Item Number_Tag",
    "_umi_duct_ITEM_NUMBER",
]

number_families = {
    "straigth",
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

do_not_tag_families = {
    "spiral duct",
    "duct spiral",
    "flex duct",
    "duct flexible",
    "boot tap wdamper",
    "gored elbow",
    "boot saddle tap",
    "coupling",
    "boot tap - wdamper",
    "rect volume damper",
    "access panel",
    "access door",
    "manbars",
    "canvas",
    "fire damper - type a",
    "fire damper - type b",
    "fire damper - type c",
    "fire damper - type cr",
    "smoke fire damper - type cr",
    "smoke fire damper - type csr",
}

values_to_skip = {
    "0",
    "skip",
}

tag_symbol = None
target_tag_name = None
for candidate in tag_names:
    tag_symbol = _find_tag_symbol(doc, candidate)
    if tag_symbol:
        target_tag_name = candidate
        break

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

fam_name = (
    tag_symbol.Family.Name if tag_symbol and tag_symbol.Family else "").strip()

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

t = Transaction(doc, "Tag Item Number")
t.Start()
try:
    # Tag fabrication ducts (exclude certain families)
    for elem in fab_ducts:
        try:
            # Get family name
            fam_param = elem.LookupParameter("Family")
            if not fam_param:
                continue

            fam_value = fam_param.AsValueString()
            if not fam_value:
                continue

            fam_lower = fam_value.strip().lower()

            # Skip families in the do_not_tag list
            if any(skip_fam in fam_lower for skip_fam in do_not_tag_families):
                continue
        except Exception:
            continue

        # Skip elements with disallowed item number values
        try:
            item_param = elem.LookupParameter("Item Number")
            if item_param:
                item_value = item_param.AsString()
                if item_value is None:
                    item_value = item_param.AsValueString()
                if item_value is None:
                    if item_param.StorageType == StorageType.Integer:
                        item_value = str(item_param.AsInteger())
                    elif item_param.StorageType == StorageType.Double:
                        item_value = str(item_param.AsDouble())
                item_value = (item_value or "").strip().lower()
                if item_value in values_to_skip:
                    continue
        except Exception:
            pass

        # Check if already tagged with our tag family
        if elem.Id in already_tagged_ids:
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
    "## Summary: placed {}, already tagged {}, failed {}".format(
        len(placed),
        len(already_tagged),
        len(failed),
    )
)
