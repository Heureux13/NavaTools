# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from revit_element import RevitElement
from revit_duct import (
    RevitDuct,
    JointSize,
    CONNECTOR_THRESHOLDS,
    DEFAULT_SHORT_THRESHOLD_IN,
)
from revit_output import print_disclaimer
from pyrevit import revit, script
from Autodesk.Revit.DB import *

# Button info
# ===================================================
__title__ = "Short"
__doc__ = """
Selects tagless short duct
"""

# Variables
# ==================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
view = revit.active_view
output = script.get_output()

family_annotations = {
    "_umi_length",
    "-fabduct_lenght",
}


def _build_attached_annotation_families_map(doc, view):
    """Map tagged element id -> set of attached tag family names (lowercase)."""
    by_element = {}
    view_id = view.Id
    tags = list(
        FilteredElementCollector(doc)
        .OfClass(IndependentTag)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    for tag in tags:
        if getattr(tag, "OwnerViewId", None) != view_id:
            continue

        try:
            getter = getattr(tag, "GetTaggedLocalElementIds", None)
            tagged_ids = []
            raw_tagged_ids = getter() if callable(getter) else None
            if raw_tagged_ids:
                enumerator = getattr(raw_tagged_ids, "GetEnumerator", None)
                if callable(enumerator):
                    it = enumerator()
                    while it.MoveNext():  # type: ignore[attr-defined]
                        # type: ignore[attr-defined]
                        tagged_ids.append(it.Current)
        except Exception:
            tagged_ids = []
            tagged_id = getattr(tag, "TaggedLocalElementId", None)
            if tagged_id:
                tagged_ids = [tagged_id]

        if not tagged_ids:
            continue

        tag_type = doc.GetElement(tag.GetTypeId())
        family_name = ""
        if tag_type and hasattr(tag_type, "Family") and tag_type.Family:
            family_name = (tag_type.Family.Name or "").strip().lower()
        if not family_name:
            continue

        for tagged_id in tagged_ids:
            if not tagged_id:
                continue
            el_id = tagged_id.IntegerValue
            by_element.setdefault(el_id, set()).add(family_name)

    return by_element


def _is_missing_required_family_annotation(duct, attached_annotations, required_annotations):
    duct_ann = attached_annotations.get(duct.element.Id.IntegerValue, set())
    return not bool(duct_ann.intersection(required_annotations))


def _is_half_inch_under_short_limit(duct):
    conn0 = (duct.connector_0_type or "").strip()
    conn1 = (duct.connector_1_type or "").strip()
    if conn0 != conn1 or duct.length is None:
        return False

    key = (duct.family, conn0)
    threshold = CONNECTOR_THRESHOLDS.get(key, DEFAULT_SHORT_THRESHOLD_IN)
    return duct.length <= (threshold - 0.5)

# Main Code
# ==================================================


# Get all ducts
ducts = RevitDuct.all(doc, view)
required_annotations = {name.strip().lower() for name in family_annotations}
attached_annotations = _build_attached_annotation_families_map(doc, view)

# Filter down to short joints at least 0.5" below threshold missing required annotations
fil_ducts = [
    d for d in ducts
    if (
        d.joint_size == JointSize.SHORT
        and _is_half_inch_under_short_limit(d)
        and _is_missing_required_family_annotation(
            d,
            attached_annotations,
            required_annotations,
        )
    )
]

# Start of select / print loop
if fil_ducts:

    # Select filtered dcuts
    RevitElement.select_many(uidoc, fil_ducts)
    output.print_md("# Selected {} short joints".format(len(fil_ducts)))
    output.print_md("---")

    # Individutal duct and selected properties
    for i, fil in enumerate(fil_ducts, start=1):
        output.print_md(
            '### No: {:03} | ID: {} | Length: {:06.2f}" | Size: {} | Connectors: 1 = {}, 2 = {}'.format(
                i,
                output.linkify(fil.element.Id),
                fil.length,
                fil.size,
                fil.connector_0_type,
                fil.connector_1_type,
            )
        )

    # Total count
    element_ids = [d.element.Id for d in fil_ducts]
    output.print_md(
        "# Total elements {}, {}".format(
            len(element_ids), output.linkify(element_ids))
    )

    # Final print statements
    print_disclaimer(output)
else:
    output.print_md("## No short joints selected")
