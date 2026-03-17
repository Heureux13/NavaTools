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

# Button info
# ==================================================
__title__ = "Remove Annotations"
__doc__ = """
Removes annotations on selected items
"""

# Variables
# ==================================================
uidoc = __revit__.ActiveUIDocument
doc = revit.doc


def is_annotation_element(element):
    """Return True when element is an annotation-like dependent element."""
    if not element:
        return False

    if isinstance(element, DB.IndependentTag):
        return True

    mra_type = getattr(DB, "MultiReferenceAnnotation", None)
    if mra_type and isinstance(element, mra_type):
        return True

    cat = element.Category
    return bool(cat and cat.CategoryType == DB.CategoryType.Annotation)


def get_annotation_dependents(host_element):
    """Collect annotation dependents attached to the host element."""
    if not host_element:
        return set()

    dependent_ids = set()
    try:
        raw_dep_ids = host_element.GetDependentElements(None)
    except Exception:
        raw_dep_ids = []

    for dep_id in raw_dep_ids or []:
        dep_elem = doc.GetElement(dep_id)
        if is_annotation_element(dep_elem):
            dependent_ids.add(dep_id)

    return dependent_ids


selected_ids = list(uidoc.Selection.GetElementIds())
if not selected_ids:
    script.exit()

host_elements = [doc.GetElement(eid) for eid in selected_ids]
annotation_ids = set()
for host in host_elements:
    annotation_ids.update(get_annotation_dependents(host))

if not annotation_ids:
    script.exit()

with revit.Transaction("Remove annotations from selected elements"):
    for ann_id in annotation_ids:
        doc.Delete(ann_id)
