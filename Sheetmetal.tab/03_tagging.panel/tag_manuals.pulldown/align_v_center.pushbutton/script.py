# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import DB, forms, revit

# Button info
# ======================================================================
__title__ = 'Vertical Center Align'
__doc__ = '''
Align selected annotations by vertical center line and
space them with a fixed 1/32" vertical gap.
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
        "top": max_pt.Y,
        "height": height,
        "width": width,
    }


def set_angle_zero_if_possible(element):
    """Set instance Angle parameter to 0 when writable."""
    try:
        angle_param = element.LookupParameter('Angle')
    except Exception:
        angle_param = None

    if not angle_param or angle_param.IsReadOnly:
        return

    try:
        if angle_param.StorageType == DB.StorageType.Double:
            angle_param.Set(0.0)
        elif angle_param.StorageType == DB.StorageType.Integer:
            angle_param.Set(0)
    except Exception:
        pass


def get_element_id_value(element_id):
    """Return a stable integer id across Revit versions."""
    if element_id is None:
        return None
    try:
        return element_id.IntegerValue
    except Exception:
        try:
            return element_id.Value
        except Exception:
            return None


def read_leader_state(element):
    """Read whether leader is enabled, using parameter first then API fallback."""
    try:
        leader_param = element.LookupParameter('Leader Line')
    except Exception:
        leader_param = None

    if leader_param and not leader_param.IsReadOnly:
        try:
            if leader_param.StorageType == DB.StorageType.Integer:
                return ('param_int', leader_param.AsInteger() != 0)
            if leader_param.StorageType == DB.StorageType.String:
                raw = leader_param.AsString() or leader_param.AsValueString() or ''
                return ('param_str', raw.strip().lower() in ('yes', 'true', '1'))
        except Exception:
            pass

    try:
        if isinstance(element, DB.IndependentTag):
            return ('has_leader', bool(element.HasLeader))
    except Exception:
        pass

    return (None, None)


def set_leader_state(element, mode, enabled):
    """Set leader enabled/disabled for supported annotation elements."""
    if mode is None:
        return

    try:
        if mode == 'param_int':
            leader_param = element.LookupParameter('Leader Line')
            if leader_param and not leader_param.IsReadOnly:
                leader_param.Set(1 if enabled else 0)
            return

        if mode == 'param_str':
            leader_param = element.LookupParameter('Leader Line')
            if leader_param and not leader_param.IsReadOnly:
                leader_param.Set('Yes' if enabled else 'No')
            return

        if mode == 'has_leader' and isinstance(element, DB.IndependentTag):
            element.HasLeader = bool(enabled)
            return
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

with revit.Transaction('Set Annotation Angle To Zero'):
    leader_states = {}
    for element in annotation_elements:
        set_angle_zero_if_possible(element)

        elem_key = get_element_id_value(element.Id)
        mode, enabled = read_leader_state(element)
        if elem_key is not None and mode is not None:
            leader_states[elem_key] = (mode, enabled)
            set_leader_state(element, mode, False)

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

# Sort by current vertical position (top-to-bottom) to preserve on-screen order.
sorted_data = sorted(
    annotation_data, key=lambda d: d["center_y"], reverse=True)
anchor = sorted_data[0]
target_center_x = anchor["center_x"]

cursor_top = anchor["top"]
for data in sorted_data:
    new_center_y = cursor_top - data["height"] / 2.0
    data["target_center_y"] = new_center_y
    cursor_top = new_center_y - \
        data["height"] / 2.0 - ONE_THIRTY_SECOND_INCH_FT

processed_count = 0
move_failures = 0

with revit.Transaction('Align + Space Selected Annotations'):
    for data in sorted_data:
        dx = target_center_x - data["center_x"]
        dy = data["target_center_y"] - data["center_y"]

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

    for element in annotation_elements:
        elem_key = get_element_id_value(element.Id)
        if elem_key is None or elem_key not in leader_states:
            continue
        mode, enabled = leader_states[elem_key]
        set_leader_state(element, mode, enabled)
