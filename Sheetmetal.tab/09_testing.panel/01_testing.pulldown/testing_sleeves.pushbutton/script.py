# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import DB, forms, revit, script
try:
    from config.parameters_registry import PYT_NUMBER_SLEEVE, PYT_SLEEVE_VALUE  # type: ignore[reportMissingImports]
except Exception:
    from lib.config.parameters_registry import PYT_NUMBER_SLEEVE, PYT_SLEEVE_VALUE

# Button info
# ======================================================================
__title__ = 'Sleeve Testing'
__doc__ = '''
Align sleeve centerline to duct centerline.

Modes:
- Selected Pair: pick exactly 2 elements (sleeve + target), then align.
- Batch by NumberSleeve: in active view, match elements by NumberSleeve and
	align sleeve candidates to non-sleeve matches.
'''

# Variables
# ======================================================================

output = script.get_output()
uidoc = getattr(revit, 'uidoc', None)
doc = getattr(revit, 'doc', None)
view = getattr(revit, 'active_view', None)

if uidoc is None or doc is None or view is None:
    forms.alert(
        'Could not access active Revit document/view context.',
        exitscript=True,
    )

assert uidoc is not None
assert doc is not None
assert view is not None

EPS = 1e-9


def _element_id_value(element_or_id):
    try:
        return element_or_id.Id.Value
    except Exception:
        pass

    try:
        return element_or_id.Id.IntegerValue
    except Exception:
        pass

    try:
        return element_or_id.Value
    except Exception:
        return element_or_id.IntegerValue


def _param_as_string(element, name):
    param = element.LookupParameter(name)
    if not param:
        return None

    try:
        val = param.AsString()
        if val is None:
            val = param.AsValueString()
        return val
    except Exception:
        return None


def _norm(value):
    if value is None:
        return ''
    return str(value).strip().lower()


def _family_name(element):
    try:
        symbol = getattr(element, 'Symbol', None)
        family = getattr(symbol, 'Family', None)
        if family is not None and family.Name:
            return str(family.Name)
    except Exception:
        pass

    try:
        if getattr(element, 'Name', None):
            return str(element.Name)
    except Exception:
        pass

    return ''


def _is_sleeve_candidate(element):
    sleeve_val = _norm(_param_as_string(element, PYT_SLEEVE_VALUE))
    if sleeve_val == 'sleeve':
        return True

    fam_name = _norm(_family_name(element))
    if 'sleeve' in fam_name:
        return True

    cat_name = _norm(getattr(getattr(element, 'Category', None), 'Name', ''))
    return 'sleeve' in cat_name


def _get_location_point(element, active_view):
    loc = getattr(element, 'Location', None)
    if loc is not None:
        pt = getattr(loc, 'Point', None)
        if pt is not None:
            return pt

        curve = getattr(loc, 'Curve', None)
        if curve is not None:
            try:
                return curve.Evaluate(0.5, True)
            except Exception:
                pass

    try:
        bbox = element.get_BoundingBox(active_view)
        if bbox is not None:
            return (bbox.Min + bbox.Max) / 2.0
    except Exception:
        pass

    return None


def _collect_connectors(element):
    connectors = []

    def _append_from_collection(collection):
        if collection is None:
            return
        try:
            for conn in collection:
                if conn is not None:
                    connectors.append(conn)
        except Exception:
            pass

    try:
        cm = getattr(element, 'ConnectorManager', None)
        if cm is not None:
            _append_from_collection(getattr(cm, 'Connectors', None))
    except Exception:
        pass

    try:
        mep = getattr(element, 'MEPModel', None)
        cm = getattr(mep, 'ConnectorManager', None) if mep is not None else None
        if cm is not None:
            _append_from_collection(getattr(cm, 'Connectors', None))
    except Exception:
        pass

    try:
        get_connectors = getattr(element, 'GetConnectors', None)
        if get_connectors is not None:
            _append_from_collection(get_connectors())
    except Exception:
        pass

    for attr in ('PrimaryConnector', 'SecondaryConnector'):
        try:
            conn = getattr(element, attr, None)
            if conn is not None:
                connectors.append(conn)
        except Exception:
            pass

    unique = []
    seen = set()
    for conn in connectors:
        try:
            origin = conn.Origin
            key = (round(origin.X, 8), round(origin.Y, 8), round(origin.Z, 8))
        except Exception:
            continue

        if key in seen:
            continue
        seen.add(key)
        unique.append(conn)

    return unique


def _line_from_connectors(element):
    connectors = _collect_connectors(element)
    if len(connectors) < 2:
        return None

    max_dist_sq = -1.0
    best_a = None
    best_b = None

    for i in range(len(connectors)):
        for j in range(i + 1, len(connectors)):
            try:
                a = connectors[i].Origin
                b = connectors[j].Origin
                dx = a.X - b.X
                dy = a.Y - b.Y
                dz = a.Z - b.Z
                dist_sq = dx * dx + dy * dy + dz * dz
            except Exception:
                continue

            if dist_sq > max_dist_sq:
                max_dist_sq = dist_sq
                best_a = a
                best_b = b

    if best_a is None or best_b is None or max_dist_sq <= EPS:
        return None

    try:
        return DB.Line.CreateBound(best_a, best_b)
    except Exception:
        return None


def _get_centerline_curve(element):
    loc = getattr(element, 'Location', None)
    if loc is not None:
        curve = getattr(loc, 'Curve', None)
        if curve is not None:
            return curve

    return _line_from_connectors(element)


def _project_point_to_curve(curve, point):
    if curve is None or point is None:
        return None

    try:
        result = curve.Project(point)
        if result is None:
            return None
        return result.XYZPoint
    except Exception:
        return None


def _distance_sq(pt_a, pt_b):
    if pt_a is None or pt_b is None:
        return None
    dx = pt_a.X - pt_b.X
    dy = pt_a.Y - pt_b.Y
    dz = pt_a.Z - pt_b.Z
    return dx * dx + dy * dy + dz * dz


def _vector_is_tiny(vec):
    return abs(vec.X) <= EPS and abs(vec.Y) <= EPS and abs(vec.Z) <= EPS


def _number_key(element):
    return _norm(_param_as_string(element, PYT_NUMBER_SLEEVE))


def _category_name(element):
    category = getattr(element, 'Category', None)
    if category is None:
        return ''
    return str(category.Name or '')


def _best_target_for_sleeve(sleeve_element, candidate_targets, active_view):
    sleeve_pt = _get_location_point(sleeve_element, active_view)
    if sleeve_pt is None:
        return None

    best = None
    best_dist_sq = None

    for target in candidate_targets:
        curve = _get_centerline_curve(target)
        if curve is not None:
            target_pt = _project_point_to_curve(curve, sleeve_pt)
        else:
            target_pt = _get_location_point(target, active_view)

        dist_sq = _distance_sq(sleeve_pt, target_pt)
        if dist_sq is None:
            continue

        if best is None or dist_sq < best_dist_sq:
            best = target
            best_dist_sq = dist_sq

    return best


def _build_alignment_move(sleeve_element, target_element, active_view):
    sleeve_pt = _get_location_point(sleeve_element, active_view)
    if sleeve_pt is None:
        return None, 'sleeve has no usable location point'

    target_curve = _get_centerline_curve(target_element)
    if target_curve is not None:
        target_pt = _project_point_to_curve(target_curve, sleeve_pt)
        if target_pt is None:
            return None, 'could not project sleeve point to target centerline'
    else:
        target_pt = _get_location_point(target_element, active_view)
        if target_pt is None:
            return None, 'target has no usable centerline or location point'

    move_vec = target_pt - sleeve_pt
    if _vector_is_tiny(move_vec):
        return DB.XYZ(0.0, 0.0, 0.0), None

    return move_vec, None


def _align_one_pair(doc_obj, active_view, sleeve_element, target_element):
    move_vec, reason = _build_alignment_move(
        sleeve_element,
        target_element,
        active_view,
    )
    if reason is not None:
        return False, reason

    if move_vec is None or _vector_is_tiny(move_vec):
        return True, 'already aligned'

    try:
        DB.ElementTransformUtils.MoveElement(
            doc_obj,
            sleeve_element.Id,
            move_vec,
        )
        return True, 'moved'
    except Exception as ex:
        return False, str(ex)


def _run_selected_pair_mode():
    selected_ids = list(uidoc.Selection.GetElementIds())
    if len(selected_ids) != 2:
        forms.alert(
            'Selected Pair mode requires exactly 2 selected elements:\n'
            '1) sleeve element\n'
            '2) target duct/accessory element',
            exitscript=True,
        )

    elem_a = doc.GetElement(selected_ids[0])
    elem_b = doc.GetElement(selected_ids[1])

    if elem_a is None or elem_b is None:
        forms.alert('Could not read selected elements.', exitscript=True)

    if _is_sleeve_candidate(elem_a) and not _is_sleeve_candidate(elem_b):
        sleeve, target = elem_a, elem_b
    elif _is_sleeve_candidate(elem_b) and not _is_sleeve_candidate(elem_a):
        sleeve, target = elem_b, elem_a
    else:
        # Fallback: first selected is moved to second selected.
        sleeve, target = elem_a, elem_b

    with revit.Transaction('Align Sleeve To Target Centerline'):
        ok, message = _align_one_pair(doc, view, sleeve, target)

    output.print_md('# Selected Pair Alignment')
    output.print_md('- Sleeve ID: {}'.format(output.linkify(sleeve.Id)))
    output.print_md('- Target ID: {}'.format(output.linkify(target.Id)))
    output.print_md('- Sleeve Category: {}'.format(_category_name(sleeve)))
    output.print_md('- Target Category: {}'.format(_category_name(target)))

    if ok:
        output.print_md('- Result: {}'.format(message))
    else:
        output.print_md('- Result: failed ({})'.format(message))


def _run_batch_mode():
    visible_elements = list(
        DB.FilteredElementCollector(doc, view.Id)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    keyed = {}
    for element in visible_elements:
        key = _number_key(element)
        if not key:
            continue

        keyed.setdefault(key, []).append(element)

    moved = []
    already_aligned = []
    skipped = []

    with revit.Transaction('Align Sleeves By NumberSleeve'):
        for key, elements in sorted(keyed.items(), key=lambda kv: kv[0]):
            sleeves = [e for e in elements if _is_sleeve_candidate(e)]
            others = [e for e in elements if not _is_sleeve_candidate(e)]

            if not sleeves:
                skipped.append((key, None, 'no sleeve candidate in match group'))
                continue

            if not others:
                skipped.append((key, None, 'no target candidate in match group'))
                continue

            for sleeve in sleeves:
                target = _best_target_for_sleeve(sleeve, others, view)
                if target is None:
                    skipped.append((
                        key,
                        sleeve,
                        'could not find nearest target in match group',
                    ))
                    continue

                ok, message = _align_one_pair(doc, view, sleeve, target)
                if not ok:
                    skipped.append((
                        key,
                        sleeve,
                        'failed to move: {}'.format(message),
                    ))
                    continue

                if message == 'already aligned':
                    already_aligned.append((key, sleeve, target))
                else:
                    moved.append((key, sleeve, target))

    output.print_md('# Batch Alignment By NumberSleeve')
    output.print_md('- Match groups found: {}'.format(len(keyed)))
    output.print_md('- Sleeves moved: {}'.format(len(moved)))
    output.print_md('- Already aligned: {}'.format(len(already_aligned)))
    output.print_md('- Skipped/failed: {}'.format(len(skipped)))

    if moved:
        output.print_md('## Moved')
        for idx, item in enumerate(moved, start=1):
            key, sleeve, target = item
            output.print_md(
                '{}. key={} | sleeve={} -> target={}'.format(
                    idx,
                    key,
                    output.linkify(sleeve.Id),
                    output.linkify(target.Id),
                )
            )

    if already_aligned:
        output.print_md('## Already Aligned')
        for idx, item in enumerate(already_aligned, start=1):
            key, sleeve, target = item
            output.print_md(
                '{}. key={} | sleeve={} -> target={}'.format(
                    idx,
                    key,
                    output.linkify(sleeve.Id),
                    output.linkify(target.Id),
                )
            )

    if skipped:
        output.print_md('## Skipped / Failed')
        for idx, item in enumerate(skipped, start=1):
            key, sleeve, reason = item
            sleeve_id = output.linkify(sleeve.Id) if sleeve is not None else '-'
            output.print_md(
                '{}. key={} | sleeve={} | reason={}'.format(
                    idx,
                    key,
                    sleeve_id,
                    reason,
                )
            )


mode_options = [
    'Selected Pair (Now)',
    'Batch by NumberSleeve (Active View)',
]

mode = forms.CommandSwitchWindow.show(
    mode_options,
    message='Choose sleeve alignment mode:',
)

if not mode:
    script.exit()

if mode == 'Selected Pair (Now)':
    _run_selected_pair_mode()
else:
    _run_batch_mode()
