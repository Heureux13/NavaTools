# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from importlib import import_module
import re

from pyrevit import forms, revit, script
from Autodesk.Revit.DB import (
    BuiltInCategory,
    ElementId,
    FilteredElementCollector,
    RevitLinkInstance,
    StorageType,
    View,
)


# Button info
# ======================================================================
__title__ = 'Section Weights Sum'
__doc__ = '''
Pick one or more views, collect fabrication ductwork + pipework,
sum _UMI_PYT_WeightPerFoot per view,
multiply by 8,
and write that value to _UMI_PYT_WeightSection on each view.
'''

SECTION_WEIGHT_MULTIPLIER = 8.0


# Variables
# ======================================================================
doc = getattr(revit, 'doc', None)
uidoc = getattr(revit, 'uidoc', None)
if doc is None or uidoc is None:
    revit_host = globals().get('__revit__')
    if revit_host is not None:
        uidoc = revit_host.ActiveUIDocument
        if uidoc is not None:
            doc = uidoc.Document

output = script.get_output()

try:
    registry = import_module('config.parameters_registry')
    PYT_WEIGHT_PER_FOOT = getattr(
        registry, 'PYT_WEIGHT_PER_FOOT', '_UMI_PYT_WeightPerFoot')
    PYT_WEIGHT_SECTION = getattr(
        registry, 'PYT_WEIGHT_SECTION', '_UMI_PYT_WeightSection')
except Exception:
    PYT_WEIGHT_PER_FOOT = '_UMI_PYT_WeightPerFoot'
    PYT_WEIGHT_SECTION = '_UMI_PYT_WeightSection'


def get_element_id_value(element_id):
    if element_id is None:
        return None
    try:
        return int(element_id.IntegerValue)
    except Exception:
        return int(element_id.Value)


def get_selected_views(document, ui_document):
    if ui_document is None:
        return []

    selected_ids = ui_document.Selection.GetElementIds()
    views = []

    for element_id in selected_ids:
        element = document.GetElement(element_id)
        if isinstance(element, View) and not element.IsTemplate:
            views.append(element)

    if views:
        return views

    picked = forms.select_views(
        title='Select view(s) for section-weight sum',
        multiple=True,
        filterfunc=lambda v: not v.IsTemplate
    )
    if not picked:
        return []

    if isinstance(picked, list):
        return picked
    return [picked]


def make_collector(document_obj, view_id=None):
    try:
        if view_id is None:
            return FilteredElementCollector(*[document_obj])
        return FilteredElementCollector(*[document_obj, view_id])
    except Exception:
        return None


def collect_fabrication_elements_in_view(document, source_view):
    duct_map = {}
    pipe_map = {}

    duct_collector = make_collector(document, source_view.Id)
    pipe_collector = make_collector(document, source_view.Id)

    if duct_collector is None or pipe_collector is None:
        return [], []

    for duct in (duct_collector
                 .OfCategoryId(ElementId(BuiltInCategory.OST_FabricationDuctwork))
                 .WhereElementIsNotElementType()
                 .ToElements()):
        duct_map[get_element_id_value(duct.Id)] = duct

    for pipe in (pipe_collector
                 .OfCategoryId(ElementId(BuiltInCategory.OST_FabricationPipework))
                 .WhereElementIsNotElementType()
                 .ToElements()):
        pipe_map[get_element_id_value(pipe.Id)] = pipe

    return list(duct_map.values()), list(pipe_map.values())


def get_world_aabb_from_bbox(bbox, extra_transform=None):
    if bbox is None:
        return None

    try:
        min_pt = bbox.Min
        max_pt = bbox.Max
    except Exception:
        return None

    bbox_transform = getattr(bbox, 'Transform', None)
    points = [min_pt, max_pt]

    xs = []
    ys = []
    zs = []

    for p in points:
        if bbox_transform is not None:
            p = bbox_transform.OfPoint(p)
        if extra_transform is not None:
            p = extra_transform.OfPoint(p)

        xs.append(p.X)
        ys.append(p.Y)
        zs.append(p.Z)

    return (min(xs), min(ys), min(zs), max(xs), max(ys), max(zs))


def aabb_intersects(a, b):
    if a is None or b is None:
        return False
    return not (
        a[3] < b[0] or a[0] > b[3] or
        a[4] < b[1] or a[1] > b[4] or
        a[5] < b[2] or a[2] > b[5]
    )


def collect_linked_fabrication_pipes_in_view(document, source_view):
    linked_pipe_map = {}

    view_crop_aabb = get_world_aabb_from_bbox(
        getattr(source_view, 'CropBox', None))

    link_instances_collector = make_collector(document, source_view.Id)
    if link_instances_collector is None:
        return []

    link_instances = (link_instances_collector
                      .OfClass(RevitLinkInstance)
                      .WhereElementIsNotElementType()
                      .ToElements())

    for link_instance in link_instances:
        try:
            get_link_doc = getattr(link_instance, 'GetLinkDocument', None)
            link_doc = get_link_doc() if callable(get_link_doc) else None
        except Exception:
            link_doc = None

        if link_doc is None:
            continue

        try:
            get_total_transform = getattr(
                link_instance, 'GetTotalTransform', None)
            link_transform = get_total_transform() if callable(get_total_transform) else None
        except Exception:
            link_transform = None

        link_collector = make_collector(link_doc)
        if link_collector is None:
            continue
        for pipe in (link_collector
                     .OfCategoryId(ElementId(BuiltInCategory.OST_FabricationPipework))
                     .WhereElementIsNotElementType()
                     .ToElements()):
            bbox_getter = getattr(pipe, 'get_BoundingBox', None)
            pipe_bbox = bbox_getter(None) if callable(bbox_getter) else None
            pipe_world_aabb = get_world_aabb_from_bbox(
                pipe_bbox, link_transform)

            # If crop is available, include only pipes intersecting host view crop region.
            if view_crop_aabb is not None and not aabb_intersects(pipe_world_aabb, view_crop_aabb):
                continue

            link_id = get_element_id_value(link_instance.Id)
            pipe_id = get_element_id_value(pipe.Id)
            if link_id is None or pipe_id is None:
                continue

            linked_pipe_map[(link_id, pipe_id)] = pipe

    return list(linked_pipe_map.values())


def lookup_parameter_case_insensitive(element, parameter_name):
    target = (parameter_name or '').strip().lower()
    if not target or element is None:
        return None

    try:
        direct = element.LookupParameter(parameter_name)
        if direct:
            return direct
    except Exception:
        pass

    for parameter in element.Parameters:
        try:
            definition = parameter.Definition
            name = definition.Name if definition else None
            if name and name.strip().lower() == target:
                return parameter
        except Exception:
            continue
    return None


def read_numeric_parameter_value(parameter):
    def _parse_first_float(text_value):
        if not text_value:
            return None

        cleaned = text_value.replace(',', '')
        match = re.search(r'[-+]?\d*\.?\d+', cleaned)
        if not match:
            return None

        try:
            return float(match.group(0))
        except Exception:
            return None

    if not parameter:
        return None

    try:
        # Shared text parameters should be parsed from raw/display text first.
        raw_text = parameter.AsString()
        parsed_text = _parse_first_float(raw_text)
        if parsed_text is not None:
            return parsed_text

        display_text = parameter.AsValueString()
        parsed_display = _parse_first_float(display_text)
        if parsed_display is not None:
            return parsed_display

        if parameter.StorageType == StorageType.Double:
            return float(parameter.AsDouble())

        if parameter.StorageType == StorageType.Integer:
            return float(parameter.AsInteger())

        if parameter.StorageType == StorageType.String:
            return None
    except Exception:
        return None

    return None


def set_numeric_parameter_value(parameter, value):
    if not parameter:
        return False, 'missing parameter'

    if parameter.IsReadOnly:
        return False, 'read-only parameter'

    try:
        if parameter.StorageType == StorageType.Double:
            parameter.Set(float(value))
            return True, None

        if parameter.StorageType == StorageType.Integer:
            parameter.Set(int(round(float(value))))
            return True, None

        if parameter.StorageType == StorageType.String:
            parameter.Set('{:.3f}'.format(float(value)))
            return True, None
    except Exception as ex:
        return False, str(ex)

    return False, 'unsupported parameter storage type'


def get_parameter_owner_element(parameter):
    if not parameter:
        return None
    try:
        return parameter.Element
    except Exception:
        return None


def get_instance_owned_parameter(element, parameter_name):
    """Return parameter only when it is owned by the element instance."""
    param = lookup_parameter_case_insensitive(element, parameter_name)
    if not param:
        return None, 'missing parameter'

    owner = get_parameter_owner_element(param)
    if owner is None:
        return None, 'unknown parameter owner'

    try:
        if owner.Id.IntegerValue != element.Id.IntegerValue:
            return None, 'parameter is type-owned (would affect other views of same type)'
    except Exception:
        return None, 'could not verify parameter ownership'

    return param, None


if uidoc is None:
    output.print_md('## Could not access active Revit UI document')
    script.exit()

if doc is None:
    output.print_md('## Could not access active Revit document')
    script.exit()

selected_views = get_selected_views(doc, uidoc)
if not selected_views:
    output.print_md('## No views selected')
    script.exit()

results = []

with revit.Transaction('Set View Section Weight'):
    for selected_view in selected_views:
        fab_ducts, fab_pipes = collect_fabrication_elements_in_view(
            doc, selected_view)
        linked_fab_pipes = collect_linked_fabrication_pipes_in_view(
            doc, selected_view)
        elements = fab_ducts + fab_pipes + linked_fab_pipes

        total_value = 0.0
        total_count = len(elements)
        value_count = 0
        missing_or_invalid = 0
        write_ok = False
        write_error = None

        for element in elements:
            param = lookup_parameter_case_insensitive(
                element, PYT_WEIGHT_PER_FOOT)
            value = read_numeric_parameter_value(param)

            if value is None:
                missing_or_invalid += 1
                value = 0.0
            else:
                value_count += 1

            total_value += value

        section_weight_sum = total_value
        section_weight_value = section_weight_sum * SECTION_WEIGHT_MULTIPLIER

        if total_count > 0:
            view_weight_param, ownership_error = get_instance_owned_parameter(
                selected_view, PYT_WEIGHT_SECTION)
            if ownership_error is not None:
                write_ok = False
                write_error = ownership_error
            else:
                write_ok, write_error = set_numeric_parameter_value(
                    view_weight_param, section_weight_value)
        else:
            write_error = 'no fabrication duct/pipe elements in view'

        results.append({
            'view_name': getattr(selected_view, 'Name', '<unknown view>'),
            'duct_count': len(fab_ducts),
            'pipe_count': len(fab_pipes),
            'linked_pipe_count': len(linked_fab_pipes),
            'total_count': total_count,
            'value_count': value_count,
            'missing_or_invalid': missing_or_invalid,
            'section_weight_sum': section_weight_sum,
            'section_weight_value': section_weight_value,
            'write_ok': write_ok,
            'write_error': write_error,
        })

output.print_md('# Section Weight Sum')
output.print_md('### Views processed: {}'.format(len(results)))

for result in results:
    output.print_md('---')
    output.print_md('### View: {}'.format(result['view_name']))
    output.print_md('### Fabrication ducts: {}'.format(result['duct_count']))
    output.print_md('### Fabrication pipes: {}'.format(result['pipe_count']))
    output.print_md('### Linked fabrication pipes: {}'.format(
        result['linked_pipe_count']))
    output.print_md('### Total elements iterated: {}'.format(
        result['total_count']))
    output.print_md('### Elements with usable {}: {}'.format(
        PYT_WEIGHT_PER_FOOT, result['value_count']))
    output.print_md(
        '### Missing/invalid {}: {}'.format(PYT_WEIGHT_PER_FOOT, result['missing_or_invalid']))
    output.print_md('## Sum {}: {:.3f}'.format(
        PYT_WEIGHT_PER_FOOT, result['section_weight_sum']))
    output.print_md('## Section weight value (sum x {}): {:.3f}'.format(
        int(SECTION_WEIGHT_MULTIPLIER), result['section_weight_value']))

    if result['write_ok']:
        output.print_md(
            '## Wrote {} on view successfully'.format(PYT_WEIGHT_SECTION))
    else:
        output.print_md('## Failed writing {} on view: {}'.format(
            PYT_WEIGHT_SECTION, result['write_error'] or 'unknown error'))
