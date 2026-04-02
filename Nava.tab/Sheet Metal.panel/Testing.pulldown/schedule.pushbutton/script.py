# pyright: reportMissingImports=false
# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.DB import FilteredElementCollector, SectionType, ViewSchedule, StorageType
from pyrevit import forms, revit, script
from System.Windows.Forms import DialogResult, OpenFileDialog
import os

import clr

clr.AddReference('System.Windows.Forms')

try:
    clr.AddReference('Microsoft.VisualBasic')
    from Microsoft.VisualBasic.FileIO import FieldType, TextFieldParser
    _HAS_TEXT_FIELD_PARSER = True
except Exception:
    FieldType = None
    TextFieldParser = None
    _HAS_TEXT_FIELD_PARSER = False

try:
    from System.Runtime.InteropServices import Marshal
except Exception:
    Marshal = None

try:
    import System
    import System.Reflection as Reflection
    _INVOKE_GET = (
        Reflection.BindingFlags.GetProperty |
        Reflection.BindingFlags.Public |
        Reflection.BindingFlags.Instance |
        Reflection.BindingFlags.IgnoreCase
    )
    _INVOKE_SET = (
        Reflection.BindingFlags.SetProperty |
        Reflection.BindingFlags.Public |
        Reflection.BindingFlags.Instance |
        Reflection.BindingFlags.IgnoreCase
    )
    _INVOKE_METHOD = (
        Reflection.BindingFlags.InvokeMethod |
        Reflection.BindingFlags.Public |
        Reflection.BindingFlags.Instance |
        Reflection.BindingFlags.IgnoreCase
    )
    _HAS_REFLECTION = True
except Exception:
    _HAS_REFLECTION = False


def _com_get(obj, name):
    """Late-bound COM property GET via IDispatch reflection."""
    return obj.GetType().InvokeMember(name, _INVOKE_GET, None, obj, None)


def _com_set(obj, name, value):
    """Late-bound COM property SET via IDispatch reflection."""
    obj.GetType().InvokeMember(
        name, _INVOKE_SET, None, obj,
        System.Array[System.Object]([value])
    )


def _com_call(obj, name, *args):
    """Late-bound COM method call via IDispatch reflection."""
    arr = System.Array[System.Object](list(args)) if args else None
    return obj.GetType().InvokeMember(name, _INVOKE_METHOD, None, obj, arr)


def _com_index(obj, *indices):
    """Late-bound COM indexed property (Item) via IDispatch reflection."""
    return obj.GetType().InvokeMember(
        'Item', _INVOKE_GET, None, obj,
        System.Array[System.Object](list(indices))
    )


# Button info
# ======================================================================
__title__ = 'Populate Schedule'
__doc__ = '''
Import Excel or CSV data and populate an existing editable Revit schedule.
'''


# Variables
# ======================================================================
doc = revit.doc
active_view = revit.active_view
output = script.get_output()


class DataOption(object):
    def __init__(self, name, rows):
        self.name = name
        self.rows = rows
        self.display_name = '{} ({} row(s))'.format(name, len(rows))


class ScheduleOption(object):
    def __init__(self, schedule, detail):
        self.schedule = schedule
        self.display_name = '{} | {}'.format(schedule.Name, detail)


def to_text(value):
    try:
        return u'{}'.format(value)
    except Exception:
        return str(value)


def clean_cell_value(value):
    if value is None:
        return ''

    try:
        if isinstance(value, float) and abs(value - round(value)) < 0.0000001:
            value = int(round(value))
    except Exception:
        pass

    text = to_text(value).strip()
    return '' if text == 'None' else text


def trim_table(rows):
    last_row_index = -1
    last_col_index = -1

    for row_index, row in enumerate(rows):
        for col_index, value in enumerate(row):
            if clean_cell_value(value):
                if row_index > last_row_index:
                    last_row_index = row_index
                if col_index > last_col_index:
                    last_col_index = col_index

    if last_row_index < 0 or last_col_index < 0:
        return []

    trimmed = []
    for row in rows[:last_row_index + 1]:
        normalized = []
        for value in row[:last_col_index + 1]:
            normalized.append(clean_cell_value(value))
        while len(normalized) < (last_col_index + 1):
            normalized.append('')
        trimmed.append(normalized)

    return trimmed


def release_com_object(com_object):
    if com_object is None or Marshal is None:
        return
    try:
        Marshal.ReleaseComObject(com_object)
    except Exception:
        pass


def read_delimited_file(file_path):
    if not _HAS_TEXT_FIELD_PARSER:
        raise Exception('Delimited file reader is unavailable in this environment.')

    parser = TextFieldParser(file_path)
    parser.TextFieldType = FieldType.Delimited
    parser.SetDelimiters(',', '\t', ';')
    parser.HasFieldsEnclosedInQuotes = True

    rows = []
    try:
        while not parser.EndOfData:
            fields = parser.ReadFields()
            rows.append([clean_cell_value(field) for field in fields])
    finally:
        parser.Close()

    return trim_table(rows)


def create_excel_application():
    if not _HAS_REFLECTION:
        return None
    try:
        excel_type = System.Type.GetTypeFromProgID('Excel.Application')
        if excel_type is None:
            return None
        return System.Activator.CreateInstance(excel_type)
    except Exception:
        return None


def read_excel_sheet_rows(worksheet):
    used_range = _com_get(worksheet, 'UsedRange')
    try:
        row_count = int(_com_get(_com_get(used_range, 'Rows'), 'Count'))
        col_count = int(_com_get(_com_get(used_range, 'Columns'), 'Count'))
        rows = []

        for row_index in range(1, row_count + 1):
            row_values = []
            for col_index in range(1, col_count + 1):
                cell = _com_index(used_range, row_index, col_index)
                value = _com_get(cell, 'Value2')
                row_values.append(clean_cell_value(value))
                release_com_object(cell)
            rows.append(row_values)

        return trim_table(rows)
    finally:
        release_com_object(used_range)


def read_excel_workbook(file_path):
    excel_app = None
    workbook = None
    options = []

    try:
        excel_app = create_excel_application()
        if excel_app is None:
            raise Exception(
                'Microsoft Excel is required to read .xlsx or .xls files on this machine.')

        _com_set(excel_app, 'Visible', False)
        _com_set(excel_app, 'DisplayAlerts', False)
        workbooks = _com_get(excel_app, 'Workbooks')
        workbook = _com_call(workbooks, 'Open', file_path)
        release_com_object(workbooks)

        worksheets = _com_get(workbook, 'Worksheets')
        sheet_count = int(_com_get(worksheets, 'Count'))
        for index in range(1, sheet_count + 1):
            worksheet = _com_index(worksheets, index)
            try:
                sheet_name = clean_cell_value(_com_get(worksheet, 'Name'))
                rows = read_excel_sheet_rows(worksheet)
                if rows:
                    options.append(DataOption(sheet_name, rows))
            finally:
                release_com_object(worksheet)
        release_com_object(worksheets)
    finally:
        if workbook is not None:
            try:
                _com_call(workbook, 'Close', False)
            except Exception:
                pass
        if excel_app is not None:
            try:
                _com_call(excel_app, 'Quit')
            except Exception:
                pass

        release_com_object(workbook)
        release_com_object(excel_app)

    return options


def pick_input_file():
    dialog = OpenFileDialog()
    dialog.Title = 'Select Excel or CSV File'
    dialog.Filter = (
        'Excel and CSV files (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv|'
        'Excel Workbook (*.xlsx)|*.xlsx|'
        'Excel 97-2003 Workbook (*.xls)|*.xls|'
        'CSV Files (*.csv)|*.csv|'
        'All Files (*.*)|*.*'
    )
    dialog.Multiselect = False

    if dialog.ShowDialog() != DialogResult.OK:
        return None

    return dialog.FileName


def get_table_data(file_path):
    extension = os.path.splitext(file_path)[1].lower()

    if extension == '.csv':
        rows = read_delimited_file(file_path)
        return os.path.splitext(os.path.basename(file_path))[0], rows

    if extension not in ('.xlsx', '.xls'):
        raise Exception('Unsupported file type: {}'.format(extension))

    sheets = read_excel_workbook(file_path)
    if not sheets:
        raise Exception('The workbook does not contain any populated worksheets.')

    selected = sheets[0]
    if len(sheets) > 1:
        selected = forms.SelectFromList.show(
            sheets,
            name_attr='display_name',
            multiselect=False,
            title='Select Worksheet'
        )
        if not selected:
            script.exit()

    return selected.name, selected.rows


def collect_schedules(document):
    schedules = []
    for schedule in FilteredElementCollector(document).OfClass(ViewSchedule):
        try:
            if schedule.IsTemplate:
                continue
        except Exception:
            pass

        try:
            definition = schedule.Definition
            detail_parts = []
            if definition.IsKeySchedule:
                detail_parts.append('Key Schedule')
            else:
                detail_parts.append('Schedule')
            if is_body_editable(schedule):
                detail_parts.append('Editable')
            else:
                detail_parts.append('Read Only')
            schedules.append(ScheduleOption(schedule, ' | '.join(detail_parts)))
        except Exception:
            pass
    return sorted(schedules, key=lambda option: option.display_name.lower())


def get_body_section(schedule):
    return schedule.GetTableData().GetSectionData(SectionType.Body)


def is_body_editable(schedule):
    try:
        body = get_body_section(schedule)
        first_row = body.FirstRowNumber
        if body.CanInsertRow(first_row):
            return True
        if body.NumberOfRows > 0 and body.CanRemoveRow(body.LastRowNumber):
            return True
    except Exception:
        pass
    return False


def pick_schedule(document, current_view):
    if isinstance(current_view, ViewSchedule):
        return current_view

    schedules = collect_schedules(document)
    if not schedules:
        return None

    selected = forms.SelectFromList.show(
        schedules,
        name_attr='display_name',
        multiselect=False,
        title='Select Schedule to Populate'
    )
    if not selected:
        return None
    return selected.schedule


def normalize_header(text):
    return clean_cell_value(text).strip().lower()


def get_schedule_headers(schedule):
    definition = schedule.Definition
    headers = []
    for field_id in definition.GetFieldOrder():
        field = definition.GetField(field_id)
        headers.append(clean_cell_value(field.ColumnHeading))
    return headers


def build_column_map(schedule_headers, source_headers):
    source_lookup = {}
    for index, header in enumerate(source_headers):
        key = normalize_header(header)
        if key and key not in source_lookup:
            source_lookup[key] = index

    mapped_columns = []
    missing_headers = []
    for schedule_col_index, schedule_header in enumerate(schedule_headers):
        key = normalize_header(schedule_header)
        if key in source_lookup:
            mapped_columns.append((schedule_col_index, source_lookup[key], schedule_header))
        else:
            missing_headers.append(schedule_header)

    return mapped_columns, missing_headers


def get_source_header_index(source_headers, expected_name):
    """Return index of source header by normalized name."""
    normalized_headers = [normalize_header(h) for h in source_headers]
    normalized_expected = normalize_header(expected_name)
    if normalized_expected in normalized_headers:
        return normalized_headers.index(normalized_expected)
    return None


def get_schedule_field_map(schedule):
    """Returns a dict of column_index -> ScheduleField for all fields."""
    definition = schedule.Definition
    field_map = {}
    for col_index, field_id in enumerate(definition.GetFieldOrder()):
        field_map[col_index] = definition.GetField(field_id)
    return field_map


def get_field_lookup_names(field):
    """Return candidate parameter names for LookupParameter fallback."""
    names = []
    try:
        internal_name = clean_cell_value(field.GetName())
        if internal_name:
            names.append(internal_name)
    except Exception:
        pass

    heading_name = clean_cell_value(field.ColumnHeading)
    if heading_name and heading_name not in names:
        names.append(heading_name)

    return names


def get_param_target_and_param(element, param_id, fallback_names):
    """Return (target_element, parameter) checking instance first, then type."""
    param = None

    try:
        param = element.get_Parameter(param_id)
    except Exception:
        param = None

    if param is None:
        for fallback_name in fallback_names:
            try:
                param = element.LookupParameter(fallback_name)
            except Exception:
                param = None
            if param is not None:
                break

    if param is not None:
        return element, param

    try:
        type_id = element.GetTypeId()
    except Exception:
        type_id = None

    if type_id is None or type_id.IntegerValue < 0:
        return None, None

    element_type = doc.GetElement(type_id)
    if element_type is None:
        return None, None

    try:
        param = element_type.get_Parameter(param_id)
    except Exception:
        param = None

    if param is None:
        for fallback_name in fallback_names:
            try:
                param = element_type.LookupParameter(fallback_name)
            except Exception:
                param = None
            if param is not None:
                break

    if param is not None:
        return element_type, param

    return None, None


def build_element_label_map(schedule, label_field):
    """Returns a dict of lowercase label value -> Element for elements in the schedule."""
    label_param_id = label_field.ParameterId
    fallback_names = get_field_lookup_names(label_field)
    element_map = {}
    for element in FilteredElementCollector(doc, schedule.Id).WhereElementIsNotElementType():
        try:
            _, param = get_param_target_and_param(element, label_param_id, fallback_names)
            if param is not None:
                label = clean_cell_value(param.AsString() or param.AsValueString())
                if label:
                    key = label.lower()
                    if key not in element_map:
                        element_map[key] = element
        except Exception:
            pass
    return element_map


def collect_source_label_samples(data_rows, label_source_col, max_count):
    samples = []
    for row_values in data_rows:
        label = row_values[label_source_col] if label_source_col < len(row_values) else ''
        label = clean_cell_value(label)
        if label:
            samples.append(label)
            if len(samples) >= max_count:
                break
    return samples


def set_element_parameter(element, field, value_text):
    _, param = get_param_target_and_param(element, field.ParameterId, get_field_lookup_names(field))
    if param is None or param.IsReadOnly:
        return False
    try:
        storage = param.StorageType
        if storage == StorageType.String:
            param.Set(value_text)
        elif storage == StorageType.Integer:
            param.Set(int(value_text)) if value_text else param.Set(0)
        elif storage == StorageType.Double:
            param.Set(float(value_text)) if value_text else param.Set(0.0)
        else:
            return False
        return True
    except Exception:
        return False


def get_body_label_map(schedule, label_schedule_col):
    """Returns a dict of lowercase label text -> row number for all body rows."""
    body = get_body_section(schedule)
    label_map = {}
    abs_col = body.FirstColumnNumber + label_schedule_col
    for row in range(body.FirstRowNumber, body.LastRowNumber + 1):
        try:
            label = clean_cell_value(schedule.GetCellText(SectionType.Body, row, abs_col))
            if label:
                label_map[label.lower()] = row
        except Exception:
            pass
    return label_map


def update_schedule_rows_by_label(schedule, data_rows, mapped_columns, label_schedule_col, label_source_col):
    field_map = get_schedule_field_map(schedule)
    label_field = field_map.get(label_schedule_col)
    element_map = build_element_label_map(schedule, label_field)

    updated = 0
    empty_label = 0
    unmatched = 0
    for row_values in data_rows:
        label = row_values[label_source_col] if label_source_col < len(row_values) else ''
        label = clean_cell_value(label)
        if not label:
            empty_label += 1
            continue
        element = element_map.get(label.lower())
        if element is None:
            unmatched += 1
            continue
        for schedule_col_index, source_col_index, _ in mapped_columns:
            if schedule_col_index == label_schedule_col:
                continue
            field = field_map.get(schedule_col_index)
            if field is None:
                continue
            cell_text = row_values[source_col_index] if source_col_index < len(row_values) else ''
            set_element_parameter(element, field, cell_text)
        updated += 1
    return updated, empty_label, unmatched, element_map


def get_import_rows(table_rows):
    if len(table_rows) < 2:
        raise Exception('The source data must include a header row and at least one data row.')

    headers = [clean_cell_value(value) for value in table_rows[0]]
    data_rows = []
    for row in table_rows[1:]:
        normalized = [clean_cell_value(value) for value in row]
        if any(normalized):
            data_rows.append(normalized)

    if not data_rows:
        raise Exception('No populated data rows were found below the header row.')

    return headers, data_rows


file_path = pick_input_file()
if not file_path:
    script.exit()

worksheet_name, table_rows = get_table_data(file_path)
if not table_rows:
    output.print_md('No data found in the selected file.')
    script.exit()

schedule = pick_schedule(doc, active_view)
if schedule is None:
    script.exit()

source_headers, data_rows = get_import_rows(table_rows)
schedule_headers = get_schedule_headers(schedule)

if not schedule_headers:
    output.print_md('Could not read the schedule column headers.')
    script.exit()

mapped_columns, missing_headers = build_column_map(schedule_headers, source_headers)
if not mapped_columns:
    output.print_md('**No source columns match the schedule headers.**')
    output.print_md('**Schedule headers:** {}'.format(', '.join(['`{}`'.format(h) for h in schedule_headers])))
    output.print_md('**Excel headers:** {}'.format(', '.join(['`{}`'.format(h) for h in source_headers])))
    script.exit()

label_schedule_col = None
for sched_col, src_col, header_name in mapped_columns:
    if normalize_header(header_name) == 'label':
        label_schedule_col = sched_col
        break

label_source_col = get_source_header_index(source_headers, 'Label')

if label_schedule_col is None or label_source_col is None:
    output.print_md('No Label column found in both schedule and source headers - cannot match rows by Label.')
    output.print_md('**Mapped columns:** {}'.format(', '.join([item[2] for item in mapped_columns])))
    output.print_md('**Source headers:** {}'.format(', '.join(source_headers)))
    script.exit()

updated_count = 0
empty_label_count = 0
unmatched_count = 0
schedule_label_count = 0
source_label_samples = []
schedule_label_samples = []

with revit.Transaction('Populate Schedule From Excel'):
    label_field = get_schedule_field_map(schedule).get(label_schedule_col)
    if label_field is not None:
        output.print_md('- Label lookup names used: {}'.format(', '.join(get_field_lookup_names(label_field))))
    updated_count, empty_label_count, unmatched_count, schedule_label_map = update_schedule_rows_by_label(
        schedule, data_rows, mapped_columns, label_schedule_col, label_source_col)
    schedule.RefreshData()

schedule_label_count = len(schedule_label_map)
source_label_samples = collect_source_label_samples(data_rows, label_source_col, 10)
schedule_label_samples = sorted(schedule_label_map.keys())[:10]

output.print_md('# Updated schedule')
output.print_md('- Schedule: {}'.format(output.linkify(schedule.Id)))
output.print_md('- Source: {}'.format(file_path))
output.print_md('- Worksheet: {}'.format(worksheet_name))
output.print_md('- Rows updated: {}'.format(updated_count))
output.print_md('- Source rows with empty Label: {}'.format(empty_label_count))
output.print_md('- Source rows with Label but no schedule match: {}'.format(unmatched_count))
output.print_md('- Excel rows with no matching label: {}'.format(empty_label_count + unmatched_count))
output.print_md('- Unique schedule labels detected: {}'.format(schedule_label_count))
output.print_md('- Label source column: {}'.format(source_headers[label_source_col]))
output.print_md('- Mapped columns: {}'.format(', '.join([item[2] for item in mapped_columns])))
if source_label_samples:
    output.print_md('- Sample source Label values: {}'.format(', '.join(source_label_samples)))
if schedule_label_samples:
    output.print_md('- Sample schedule Label values: {}'.format(', '.join(schedule_label_samples)))

if missing_headers:
    output.print_md('- Schedule columns not in Excel: {}'.format(', '.join(missing_headers)))
