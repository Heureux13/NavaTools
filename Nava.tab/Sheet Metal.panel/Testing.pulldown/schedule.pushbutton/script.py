# pyright: reportMissingImports=false
# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.DB import FilteredElementCollector, SectionType, ViewSchedule
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
    clr.AddReference('Microsoft.Office.Interop.Excel')
    from Microsoft.Office.Interop import Excel
    _HAS_EXCEL_INTEROP = True
except Exception:
    Excel = None
    _HAS_EXCEL_INTEROP = False

try:
    from System.Runtime.InteropServices import Marshal
except Exception:
    Marshal = None


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
    if not _HAS_EXCEL_INTEROP:
        return None
    try:
        return Excel.ApplicationClass()
    except Exception:
        return Excel.Application()


def read_excel_sheet_rows(worksheet):
    used_range = worksheet.UsedRange
    try:
        row_count = used_range.Rows.Count
        col_count = used_range.Columns.Count
        rows = []

        for row_index in range(1, row_count + 1):
            row_values = []
            for col_index in range(1, col_count + 1):
                value = used_range.Cells[row_index, col_index].Value2
                row_values.append(clean_cell_value(value))
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

        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        workbook = excel_app.Workbooks.Open(file_path)

        for index in range(1, workbook.Worksheets.Count + 1):
            worksheet = workbook.Worksheets[index]
            try:
                rows = read_excel_sheet_rows(worksheet)
                if rows:
                    options.append(DataOption(worksheet.Name, rows))
            finally:
                release_com_object(worksheet)
    finally:
        if workbook is not None:
            try:
                workbook.Close(False)
            except Exception:
                pass
        if excel_app is not None:
            try:
                excel_app.Quit()
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
    header_section = schedule.GetTableData().GetSectionData(SectionType.Header)
    body_section = get_body_section(schedule)
    headers = []

    header_row = header_section.LastRowNumber
    first_col = body_section.FirstColumnNumber
    last_col = body_section.LastColumnNumber

    for col in range(first_col, last_col + 1):
        header_text = clean_cell_value(schedule.GetCellText(SectionType.Header, header_row, col))
        headers.append(header_text)

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


def clear_schedule_body(schedule):
    removed = 0
    body = get_body_section(schedule)
    for row in range(body.LastRowNumber, body.FirstRowNumber - 1, -1):
        if body.CanRemoveRow(row):
            body.RemoveRow(row)
            removed += 1
    return removed


def insert_schedule_rows(schedule, data_rows, mapped_columns):
    body = get_body_section(schedule)
    inserted = 0
    for row_values in data_rows:
        insert_index = body.LastRowNumber + 1 if body.NumberOfRows > 0 else body.FirstRowNumber
        if not body.CanInsertRow(insert_index):
            raise Exception('Revit does not allow inserting rows into this schedule body.')

        body.InsertRow(insert_index)
        body = get_body_section(schedule)

        for schedule_col_index, source_col_index, _ in mapped_columns:
            cell_text = row_values[source_col_index] if source_col_index < len(row_values) else ''
            body.SetCellText(insert_index, body.FirstColumnNumber + schedule_col_index, cell_text)
        inserted += 1

    return inserted


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
    forms.alert('No data found in the selected file.', exitscript=True)

schedule = pick_schedule(doc, active_view)
if schedule is None:
    script.exit()

if not is_body_editable(schedule):
    forms.alert(
        'This schedule body is not directly editable by the Revit API. '
        'Use a key schedule, sheet list, or another editable schedule type.',
        exitscript=True
    )

source_headers, data_rows = get_import_rows(table_rows)
schedule_headers = get_schedule_headers(schedule)

if not schedule_headers:
    forms.alert('Could not read the schedule column headers.', exitscript=True)

mapped_columns, missing_headers = build_column_map(schedule_headers, source_headers)
if not mapped_columns:
    forms.alert(
        'No source columns match the schedule headers. '
        'Make the spreadsheet header names match the schedule column names.',
        exitscript=True
    )

removed_count = 0
inserted_count = 0

with revit.Transaction('Populate Schedule From Excel'):
    removed_count = clear_schedule_body(schedule)
    inserted_count = insert_schedule_rows(schedule, data_rows, mapped_columns)
    schedule.RefreshData()

output.print_md('# Populated schedule')
output.print_md('- Schedule: {}'.format(output.linkify(schedule.Id)))
output.print_md('- Source: {}'.format(file_path))
output.print_md('- Worksheet: {}'.format(worksheet_name))
output.print_md('- Imported rows: {}'.format(inserted_count))
output.print_md('- Cleared rows: {}'.format(removed_count))
output.print_md('- Mapped columns: {}'.format(', '.join([item[2] for item in mapped_columns])))

if missing_headers:
    output.print_md('- Unmatched schedule columns: {}'.format(', '.join(missing_headers)))

if not schedule.Definition.IsKeySchedule:
    output.print_md('')
    output.print_md('This workflow is safest for key schedules and other editable schedule types.')
