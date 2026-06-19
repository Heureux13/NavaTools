# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

import math
from pyrevit import DB, forms, revit

# Button info
# ======================================================================
__title__ = 'Horizontal Center Align'
__doc__ = '''
Align selected annotations by horizontal center line and
space them with a fixed 1/32" horizontal gap.
Sets Angle to 90 before alignment.
'''

# Variables
# ======================================================================

uidoc = getattr(revit, 'uidoc', None)
doc = uidoc.Document if uidoc else getattr(revit, 'doc', None)
active_view = uidoc.ActiveView if uidoc else getattr(
    revit, 'active_view', None)

ONE_THIRTY_SECOND_INCH_FT = 1.0 / 32.0 / 12.0

if uidoc is None or doc is None or active_view is None:
    forms.alert(
        'Could not access active Revit document/view context.', exitscript=True)

assert uidoc is not None
assert doc is not None
assert active_view is not None


def is_annotation_element(element):
    """Return True when the element belongs to an annotation category."""
    if element is None:
        return False

    category = element.Category
    return bool(category and category.CategoryType == DB.CategoryType.Annotation)


def get_bbox_data(element):
    """Read view bounding-box geometry used for alignment and spacing."""
    try:
        bbox = element.get_BoundingBox(active_view)
    except Exception:
        return None

    if bbox is None:
        return None

    min_pt = bbox.Min
    max_pt = bbox.Max

    width = max_pt.X - min_pt.X
    height = max_pt.Y - min_pt.Y
    if height <= 0:
        return None

    center_x = (min_pt.X + max_pt.X) / 2.0
    center_y = (min_pt.Y + max_pt.Y) / 2.0

    return {
        "element": element,
        "center_x": center_x,
        "center_y": center_y,
        "left": min_pt.X,
        "top": max_pt.Y,
        "height": height,
        "width": width,
    }


def set_angle_ninety_if_possible(element):
    """Set instance Angle parameter to 90 degrees when writable."""
    try:
        angle_param = element.LookupParameter('Angle')
    except Exception:
        angle_param = None

    if not angle_param or angle_param.IsReadOnly:
        return

    try:
        if angle_param.StorageType == DB.StorageType.Double:
            angle_param.Set(math.pi / 2.0)
        elif angle_param.StorageType == DB.StorageType.Integer:
            angle_param.Set(90)
    except Exception:
        pass


selected_ids = list(uidoc.Selection.GetElementIds())
if len(selected_ids) < 2:
    forms.alert(
        'Select at least 2 annotations, then run again.',
        exitscript=True,
    )

annotation_elements = []

for elem_id in selected_ids:
    element = doc.GetElement(elem_id)
    if not is_annotation_element(element):
        continue

    annotation_elements.append(element)

if len(annotation_elements) < 2:
    forms.alert(
        'Need at least 2 selected annotation elements.',
        exitscript=True,
    )

with revit.Transaction('Set Annotation Angle To 90'):
    for element in annotation_elements:
        set_angle_ninety_if_possible(element)
    doc.Regenerate()

annotation_data = []
for element in annotation_elements:
    data = get_bbox_data(element)
    if data is None:
        continue
    annotation_data.append(data)

if len(annotation_data) < 2:
    forms.alert(
        'Need at least 2 selected annotations with valid view bounding boxes.',
        exitscript=True,
    )

# Keep the left-most annotation fixed as the alignment/spatial anchor.
sorted_data = sorted(annotation_data, key=lambda d: d["center_x"])
anchor = sorted_data[0]
target_center_y = anchor["center_y"]

cursor_left = anchor["left"]
for data in sorted_data:
    new_center_x = cursor_left + data["width"] / 2.0
    data["target_center_x"] = new_center_x
    cursor_left = new_center_x + \
        data["width"] / 2.0 + ONE_THIRTY_SECOND_INCH_FT

processed_count = 0
move_failures = 0

with revit.Transaction('Align + Space Selected Annotations'):
    for data in sorted_data:
        dx = data["target_center_x"] - data["center_x"]
        dy = target_center_y - data["center_y"]

        if abs(dx) < 1e-10 and abs(dy) < 1e-10:
            processed_count += 1
            continue

        try:
            DB.ElementTransformUtils.MoveElement(  # type: ignore[reportCallIssue, reportAttributeAccessIssue]
                doc,
                data["element"].Id,
                DB.XYZ(dx, dy, 0.0),  # type: ignore[reportCallIssue]
            )
            processed_count += 1
        except Exception:
            move_failures += 1
