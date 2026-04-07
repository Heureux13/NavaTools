# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Button info
# ======================================================================
__title__ = 'Create Project Parameters'
__doc__ = '''
Create project parameters from shared parameters using the project parameter map.
Map sections determine both which parameter names are processed and parameter group assignment.
'''

from Autodesk.Revit.DB import (
    BuiltInCategory,
    Transaction,
)
from pyrevit import revit, script

import os
import sys

doc = revit.doc
app = doc.Application
output = script.get_output()

SCRIPT_DIR = os.path.dirname(__file__)
PARAMETER_MAP_PATH = None
LIB_DIR = None

search_dir = SCRIPT_DIR
while True:
    candidate = os.path.join(
        search_dir, 'lib', 'constants', 'project_parameter_map.py')
    if os.path.exists(candidate):
        PARAMETER_MAP_PATH = candidate
        LIB_DIR = os.path.join(search_dir, 'lib')
        break

    parent = os.path.dirname(search_dir)
    if parent == search_dir:
        break
    search_dir = parent

if not PARAMETER_MAP_PATH or not LIB_DIR:
    output.print_md('## Error')
    output.print_md(
        '- Could not locate lib/constants/project_parameter_map.py by searching parent folders from script location.')
    script.exit()

if LIB_DIR not in sys.path:
    sys.path.append(LIB_DIR)

project_parameter_map = __import__(
    'constants.project_parameter_map', fromlist=['*'])


def _normalize_parameter_names(raw_values):
    names = []
    for value in raw_values:
        if isinstance(value, str) and value.startswith('_UMI_'):
            names.append(value)
    return sorted(set(names))


def _collect_mapped_parameter_entries():
    entries = []

    seg_fit_values = getattr(project_parameter_map,
                             'segments_and_fittings', set())
    for name in _normalize_parameter_names(seg_fit_values):
        entries.append((name, 'Segments and Fittings'))

    construction_values = getattr(project_parameter_map, 'construction', set())
    for name in _normalize_parameter_names(construction_values):
        entries.append((name, 'Construction'))

    return entries


def _get_requested_categories(document):
    requested = (
        BuiltInCategory.OST_DuctTerminal,
        BuiltInCategory.OST_FabricationDuctwork,
        BuiltInCategory.OST_MechanicalEquipment,
    )

    category_set = app.Create.NewCategorySet()
    missing = []
    for bic in requested:
        category = document.Settings.Categories.get_Item(bic)
        if category:
            category_set.Insert(category)
        else:
            missing.append(str(bic))

    return category_set, missing


def _find_shared_definition(shared_file, param_name):
    for group in shared_file.Groups:
        definition = group.Definitions.get_Item(param_name)
        if definition:
            return definition
    return None


def _get_group_id_by_label(group_name):
    db = __import__('Autodesk.Revit.DB', fromlist=['*'])
    target = group_name.strip().lower()
    tokens = [t for t in target.replace('&', 'and').split(' ') if t]

    # Revit 2022+ path: discover built-in groups and match by UI label.
    if hasattr(db, 'ParameterUtils') and hasattr(db, 'LabelUtils'):
        try:
            for group_id in db.ParameterUtils.GetAllBuiltInGroups():
                try:
                    label = db.LabelUtils.GetLabelForGroup(group_id)
                except Exception:
                    continue

                if not label:
                    continue

                normalized = label.strip().lower()
                if normalized == target:
                    return group_id, label

                if all(token in normalized for token in tokens):
                    return group_id, label
        except Exception:
            pass

    # Fallback guesses for environments where label enumeration is unavailable.
    if hasattr(db, 'GroupTypeId'):
        fallback_map = {
            'segments and fittings': ('Fabrication', 'SegmentsAndFittings'),
            'construction': ('Construction',),
        }
        for attr_name in fallback_map.get(target, ()):
            if hasattr(db.GroupTypeId, attr_name):
                return getattr(db.GroupTypeId, attr_name), attr_name

    return None, None


def _bind_instance_parameter(definition, category_set, group_id):
    binding = app.Create.NewInstanceBinding(category_set)

    # Revit API overloads differ by version; prefer the simplest compatible call.
    if group_id is not None:
        try:
            inserted = doc.ParameterBindings.Insert(
                definition, binding, group_id)
        except TypeError:
            inserted = doc.ParameterBindings.Insert(definition, binding)
    else:
        inserted = doc.ParameterBindings.Insert(definition, binding)

    if not inserted:
        if group_id is not None:
            try:
                inserted = doc.ParameterBindings.ReInsert(
                    definition, binding, group_id)
            except TypeError:
                inserted = doc.ParameterBindings.ReInsert(definition, binding)
        else:
            inserted = doc.ParameterBindings.ReInsert(definition, binding)

    return inserted


shared_file = app.OpenSharedParameterFile()
if not shared_file:
    output.print_md('## Error')
    output.print_md(
        '- No shared parameter file is configured in Revit. Configure it first and run again.')
    script.exit()

category_set, missing_categories = _get_requested_categories(doc)
if category_set.IsEmpty:
    output.print_md('## Error')
    output.print_md(
        '- None of the requested categories were found in this document.')
    script.exit()

if missing_categories:
    output.print_md('### Missing categories in this model')
    for missing in missing_categories:
        output.print_md('- {}'.format(missing))

parameter_entries = _collect_mapped_parameter_entries()
if not parameter_entries:
    output.print_md('## Error')
    output.print_md(
        '- No parameter names were found in constants.project_parameter_map.')
    script.exit()

group_resolution = {}
for group_name in ('Segments and Fittings', 'Construction'):
    group_id, group_label = _get_group_id_by_label(group_name)
    group_resolution[group_name] = (group_id, group_label)

results = {
    'found': 0,
    'missing_definition': [],
    'instance_bound': 0,
    'instance_failed': [],
}

output.print_md('### Parameter group resolution')
for group_name in ('Segments and Fittings', 'Construction'):
    group_id, group_label = group_resolution[group_name]
    if group_id is None:
        output.print_md(
            '- {}: not resolved (may appear under Other)'.format(group_name))
    else:
        output.print_md('- {}: {}'.format(group_name, group_label))

t = Transaction(doc, 'Bind BBM and PYT project parameters')
t.Start()
try:
    for param_name, group_name in parameter_entries:
        definition = _find_shared_definition(shared_file, param_name)
        if not definition:
            results['missing_definition'].append(
                '{} [{}]'.format(param_name, group_name))
            continue

        results['found'] += 1

        group_id, _ = group_resolution.get(group_name, (None, None))

        if _bind_instance_parameter(definition, category_set, group_id):
            results['instance_bound'] += 1
        else:
            results['instance_failed'].append(
                '{} [{}]'.format(param_name, group_name))

    t.Commit()
except Exception:
    t.RollBack()
    raise

output.print_md('## BBM/PYT Project Parameter Binding')
output.print_md(
    '- Total mapped names in project_parameter_map: **{}**'.format(len(parameter_entries)))
output.print_md('- Shared definitions found: **{}**'.format(results['found']))
output.print_md(
    '- Instance bindings inserted/updated: **{}**'.format(results['instance_bound']))

if results['missing_definition']:
    output.print_md('### Missing in shared parameter file ({})'.format(
        len(results['missing_definition'])))
    for name in results['missing_definition']:
        output.print_md('- {}'.format(name))

if results['instance_failed']:
    output.print_md('### Instance binding failures ({})'.format(
        len(results['instance_failed'])))
    for name in results['instance_failed']:
        output.print_md('- {}'.format(name))
