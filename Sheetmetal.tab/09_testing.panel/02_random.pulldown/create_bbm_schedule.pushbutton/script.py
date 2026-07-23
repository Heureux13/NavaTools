# pyright: reportMissingImports=false
# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from config import parameters_registry
from Autodesk.Revit.DB import (
    FilteredElementCollector, ViewSchedule, ScheduleFilter, ScheduleFilterType
)
from Autodesk.Revit.DB import ElementId
from pyrevit import revit, script
import os
import sys

# Import parameters registry
lib_path = os.path.join(os.path.dirname(__file__), '..', '..', '..', 'lib')
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)


__title__ = 'Create BBM Schedules'
__doc__ = '''
Create a schedule with all Bluebeam Map (BBM_) parameters from parameters_registry.
'''

doc = revit.doc
output = script.get_output()

# Only schedules for categories listed here will be created/replaced.
categories_to_use = [
    'Air Terminals',
    'Mechanical Equipment',
    'MEP Fabrication Ductwork',
]

category_aliases = {
    # Common typo aliases
    'mepfabricationductkwork': 'mepfabricationductwork',
}


def _normalize_category_name(text):
    """Normalize category names for robust matching."""
    return ''.join(ch for ch in text.lower() if ch.isalnum())


def get_selected_categories(document, selected_names):
    """Return requested categories from the current document by name."""
    available = {}
    checked = 0
    for category in document.Settings.Categories:
        checked += 1
        try:
            if category is None:
                continue
            key = _normalize_category_name(category.Name)
            if key not in available:
                available[key] = (category.Id.IntegerValue, category.Name)
        except Exception:
            pass

    selected = {}
    missing = []
    for raw_name in selected_names:
        key = _normalize_category_name(raw_name)
        key = category_aliases.get(key, key)
        match = available.get(key)
        if match is None:
            missing.append(raw_name)
            continue
        category_id, category_name = match
        selected[category_id] = category_name

    return selected, missing, checked


def get_bbm_parameters():
    """Extract all BBM_ parameter constants and their values from registry."""
    bbm_params = {}
    for attr_name in dir(parameters_registry):
        if attr_name.startswith('BBM_'):
            param_value = getattr(parameters_registry, attr_name, None)
            if isinstance(param_value, str):
                bbm_params[attr_name] = param_value
    return bbm_params


def schedule_exists(document, schedule_name):
    """Check if a schedule with the given name already exists."""
    for schedule in FilteredElementCollector(document).OfClass(ViewSchedule):
        try:
            if schedule.Name.lower().strip() == schedule_name.lower().strip():
                return True
        except Exception:
            pass
    return False


def get_existing_schedule_names(document):
    """Get existing schedule names in lowercase for fast membership checks."""
    names = set()
    for schedule in FilteredElementCollector(document).OfClass(ViewSchedule):
        try:
            names.add(schedule.Name.lower().strip())
        except Exception:
            pass
    return names


def get_existing_schedule_lookup(document):
    """Get existing schedules by lowercase name for quick access."""
    lookup = {}
    for schedule in FilteredElementCollector(document).OfClass(ViewSchedule):
        try:
            lookup[schedule.Name.lower().strip()] = schedule
        except Exception:
            pass
    return lookup


def get_parameter_by_name(document, param_name):
    """Get parameter element by shared parameter name."""
    try:
        from Autodesk.Revit.DB import SharedParameterElement
        for param in FilteredElementCollector(document).OfClass(SharedParameterElement):
            try:
                if param.GetName() == param_name:
                    return param
            except Exception:
                pass
    except Exception:
        pass
    return None


def get_all_valid_categories(document):
    """Get document categories likely to support scheduling."""
    valid_categories = {}
    attempted = 0
    found_count = 0

    # Iterate categories from the document instead of every OST_* enum entry.
    for category in document.Settings.Categories:
        attempted += 1
        try:
            if category is None:
                continue
            if hasattr(category, 'IsTagCategory') and category.IsTagCategory:
                continue
            if hasattr(category, 'AllowsBoundParameters') and not category.AllowsBoundParameters:
                continue

            category_id = category.Id.IntegerValue
            valid_categories[category_id] = category.Name
            found_count += 1
        except Exception:
            pass

    output.print_md(
        '  - Checked {} document categories, added {} categories'.format(attempted, found_count))
    return valid_categories


def get_schedulable_field_map(definition, document):
    """Return {field_name: SchedulableField} for a schedule definition."""
    field_map = {}
    try:
        for schedulable_field in definition.GetSchedulableFields():
            try:
                field_name = schedulable_field.GetName(document)
                if field_name and field_name not in field_map:
                    field_map[field_name] = schedulable_field
            except Exception:
                pass
    except Exception:
        pass
    return field_map


def create_bbm_schedules_for_categories(document, bbm_params, categories_dict):
    """Create schedules for specified categories and add BBM parameters."""
    created_schedules = []
    skipped_categories = []
    failed_categories = []
    existing_names = get_existing_schedule_names(document)
    existing_lookup = get_existing_schedule_lookup(document)

    with revit.Transaction('Create BBM Schedules'):
        sorted_categories = sorted(categories_dict.items(), key=lambda x: x[1])
        for idx, (category_id, category_name) in enumerate(sorted_categories, 1):
            schedule_name = '_BBM_{}'.format(category_name)

            # Check if schedule already exists
            if schedule_name.lower().strip() in existing_names:
                try:
                    existing_schedule = existing_lookup.get(
                        schedule_name.lower().strip())
                    if existing_schedule is not None:
                        document.Delete(existing_schedule.Id)
                    existing_names.discard(schedule_name.lower().strip())
                    existing_lookup.pop(schedule_name.lower().strip(), None)
                except Exception:
                    skipped_categories.append(
                        (category_name, 'Existing schedule could not be replaced'))
                    if idx % 10 == 0:
                        output.print_md('  - Progress: {} / {} categories processed'.format(
                            idx, len(sorted_categories)))
                    continue

            try:
                # Create schedule for this category
                schedule = ViewSchedule.CreateSchedule(
                    document, ElementId(category_id))
                schedule.Name = schedule_name

                definition = schedule.Definition
                added_count = 0
                failed_params = []
                schedulable_fields = get_schedulable_field_map(
                    definition, document)

                # Add BBM parameters as fields
                for param_const_name, param_name in sorted(bbm_params.items()):
                    schedulable_field = schedulable_fields.get(param_name)
                    if schedulable_field is None:
                        failed_params.append(param_name)
                        continue

                    try:
                        definition.AddField(schedulable_field)
                        added_count += 1
                    except Exception:
                        # Parameter exists but is not schedulable for this category.
                        failed_params.append(param_name)

                # Apply filter: hide rows where _UMI_BBM_Label is blank/null
                label_param_name = '_UMI_BBM_Label'
                label_field = schedulable_fields.get(label_param_name)
                if label_field is not None:
                    try:
                        added_label_field = definition.AddField(label_field)
                        schedule_filter = ScheduleFilter(
                            added_label_field.FieldId,
                            ScheduleFilterType.Contains,
                            '-'
                        )
                        definition.AddFilter(schedule_filter)
                    except Exception:
                        pass

                if added_count == 0:
                    document.Delete(schedule.Id)
                    skipped_categories.append(
                        (category_name, 'No BBM parameters schedulable for category'))
                    if idx % 10 == 0:
                        output.print_md('  - Progress: {} / {} categories processed'.format(
                            idx, len(sorted_categories)))
                    continue

                created_schedules.append({
                    'name': schedule_name,
                    'category': category_name,
                    'added': added_count,
                    'failed': failed_params
                })
                existing_names.add(schedule_name.lower().strip())
                existing_lookup[schedule_name.lower().strip()] = schedule

                if idx % 10 == 0:
                    output.print_md('  - Progress: {} / {} categories processed'.format(
                        idx, len(sorted_categories)))

            except Exception as e:
                failed_categories.append((category_name, str(e)))
                if idx % 10 == 0:
                    output.print_md('  - Progress: {} / {} categories processed'.format(
                        idx, len(sorted_categories)))

    return created_schedules, skipped_categories, failed_categories


# Main execution
bbm_parameters = get_bbm_parameters()
output.print_md('# Create BBM Schedules')
output.print_md('Found {} BBM parameters'.format(len(bbm_parameters)))

if not bbm_parameters:
    output.print_md('**No BBM parameters found in parameters_registry!**')
    script.exit()

output.print_md('\nResolving categories from categories_to_use list...')
all_categories, missing_categories, checked_categories = get_selected_categories(
    doc, categories_to_use)
output.print_md('Checked {} document categories'.format(checked_categories))
output.print_md('Requested {} categories, found {}'.format(
    len(categories_to_use), len(all_categories)))

if missing_categories:
    output.print_md('\n## Categories Not Found ({})'.format(
        len(missing_categories)))
    for name in missing_categories:
        output.print_md('- {}'.format(name))

if not all_categories:
    output.print_md('\n**None of the requested categories were found.**')
    script.exit()

created, skipped, failed = create_bbm_schedules_for_categories(
    doc, bbm_parameters, all_categories)

if created:
    output.print_md('\n## Schedules Created ({})'.format(len(created)))
    for sched in created:
        output.print_md(
            '- **{}** ({})'.format(sched['name'], sched['category']))
        output.print_md('  - Parameters added: {}'.format(sched['added']))
        if sched['failed']:
            output.print_md(
                '  - Parameters not found: {}'.format(len(sched['failed'])))

if skipped:
    output.print_md('\n## Schedules Skipped ({})'.format(len(skipped)))
    for cat_name, reason in skipped:
        output.print_md('- **{}**: {}'.format(cat_name, reason))

if failed:
    output.print_md('\n## Failed Categories ({})'.format(len(failed)))
    for cat_name, error in failed:
        output.print_md('- **{}**: {}'.format(cat_name, error))

if not created and not skipped and not failed:
    output.print_md(
        '\n**No categories found or no schedules could be created.**')
elif created:
    output.print_md(
        '\n**Success!** {} schedule(s) created.'.format(len(created)))
else:
    output.print_md('\n**No new schedules were created.**')
