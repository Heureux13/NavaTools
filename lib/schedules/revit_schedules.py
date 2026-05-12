# pyright: reportMissingImports=false, reportAttributeAccessIssue=false, reportOptionalCall=false, reportOptionalMemberAccess=false, reportGeneralTypeIssues=false
# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

import os

import clr
from Autodesk.Revit.DB import FilteredElementCollector, SectionType, ViewSchedule, StorageType
from pyrevit import forms, revit, script
from System.Windows.Forms import DialogResult, OpenFileDialog

from constants.bluebeam_map import COLUMN_MAP

try:
    from Autodesk.Revit.DB import UnitUtils, UnitTypeId, SpecTypeId
except Exception:
    UnitUtils = None
    UnitTypeId = None
    SpecTypeId = None

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


class DataOption(object):
    def __init__(self, name, rows):
        self.name = name
        self.rows = rows
        self.display_name = '{} ({} row(s))'.format(name, len(rows))


class ScheduleOption(object):
    def __init__(self, schedule, detail):
        self.schedule = schedule
        self.display_name = '{} | {}'.format(schedule.Name, detail)


class RevitSchedules(object):
    """Utilities to import tabular data and populate Revit schedules by Label."""

    def __init__(self, doc=None, active_view=None, output_obj=None, source_header_aliases=None):
        self.doc = doc or revit.doc
        self.active_view = active_view or revit.active_view
        self.output = output_obj or script.get_output()
        self.source_header_aliases = source_header_aliases or self._build_source_header_aliases()

    @staticmethod
    def _build_source_header_aliases():
        aliases = {}
        for bluebeam_col, config in COLUMN_MAP.items():
            if config['aliases']:
                revit_param = config['aliases'][0].lstrip('_')
                aliases[revit_param] = [
                    bluebeam_col.lower().replace(' ', '_'),
                    bluebeam_col.lower(),
                ]
        return aliases

    @staticmethod
    def to_text(value):
        try:
            return u'{}'.format(value)
        except Exception:
            return str(value)

    @classmethod
    def clean_cell_value(cls, value):
        if value is None:
            return ''

        try:
            if isinstance(value, float) and abs(value - round(value)) < 0.0000001:
                value = int(round(value))
        except Exception:
            pass

        text = cls.to_text(value).strip()
        return '' if text == 'None' else text

    @classmethod
    def normalize_header(cls, text):
        return cls.clean_cell_value(text).strip().lower()

    @classmethod
    def trim_table(cls, rows):
        last_row_index = -1
        last_col_index = -1

        for row_index, row in enumerate(rows):
            for col_index, value in enumerate(row):
                if cls.clean_cell_value(value):
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
                normalized.append(cls.clean_cell_value(value))
            while len(normalized) < (last_col_index + 1):
                normalized.append('')
            trimmed.append(normalized)

        return trimmed

    @staticmethod
    def _release_com_object(com_object):
        if com_object is None or Marshal is None:
            return
        try:
            Marshal.ReleaseComObject(com_object)
        except Exception:
            pass

    @staticmethod
    def _com_get(obj, name):
        return obj.GetType().InvokeMember(name, _INVOKE_GET, None, obj, None)

    @staticmethod
    def _com_set(obj, name, value):
        obj.GetType().InvokeMember(
            name, _INVOKE_SET, None, obj,
            System.Array[System.Object]([value])
        )

    @staticmethod
    def _com_call(obj, name, *args):
        arr = System.Array[System.Object](list(args)) if args else None
        return obj.GetType().InvokeMember(name, _INVOKE_METHOD, None, obj, arr)

    @staticmethod
    def _com_index(obj, *indices):
        return obj.GetType().InvokeMember(
            'Item', _INVOKE_GET, None, obj,
            System.Array[System.Object](list(indices))
        )

    @staticmethod
    def _create_excel_application():
        if not _HAS_REFLECTION:
            return None
        try:
            excel_type = System.Type.GetTypeFromProgID('Excel.Application')
            if excel_type is None:
                return None
            return System.Activator.CreateInstance(excel_type)
        except Exception:
            return None

    def read_delimited_file(self, file_path):
        if not _HAS_TEXT_FIELD_PARSER:
            raise Exception(
                'Delimited file reader is unavailable in this environment.')

        parser = TextFieldParser(file_path)
        parser.TextFieldType = FieldType.Delimited
        parser.SetDelimiters(',', '\t', ';')
        parser.HasFieldsEnclosedInQuotes = True

        rows = []
        try:
            while not parser.EndOfData:
                fields = parser.ReadFields()
                rows.append([self.clean_cell_value(field) for field in fields])
        finally:
            parser.Close()

        return self.trim_table(rows)

    def _read_excel_sheet_rows(self, worksheet):
        used_range = self._com_get(worksheet, 'UsedRange')
        try:
            row_count = int(self._com_get(
                self._com_get(used_range, 'Rows'), 'Count'))
            col_count = int(self._com_get(
                self._com_get(used_range, 'Columns'), 'Count'))
            rows = []

            for row_index in range(1, row_count + 1):
                row_values = []
                for col_index in range(1, col_count + 1):
                    cell = self._com_index(used_range, row_index, col_index)
                    value = self._com_get(cell, 'Value2')
                    row_values.append(self.clean_cell_value(value))
                    self._release_com_object(cell)
                rows.append(row_values)

            return self.trim_table(rows)
        finally:
            self._release_com_object(used_range)

    def read_excel_workbook(self, file_path):
        excel_app = None
        workbook = None
        options = []

        try:
            excel_app = self._create_excel_application()
            if excel_app is None:
                raise Exception(
                    'Microsoft Excel is required to read .xlsx or .xls files on this machine.')

            self._com_set(excel_app, 'Visible', False)
            self._com_set(excel_app, 'DisplayAlerts', False)
            workbooks = self._com_get(excel_app, 'Workbooks')
            workbook = self._com_call(workbooks, 'Open', file_path)
            self._release_com_object(workbooks)

            worksheets = self._com_get(workbook, 'Worksheets')
            sheet_count = int(self._com_get(worksheets, 'Count'))
            for index in range(1, sheet_count + 1):
                worksheet = self._com_index(worksheets, index)
                try:
                    sheet_name = self.clean_cell_value(
                        self._com_get(worksheet, 'Name'))
                    rows = self._read_excel_sheet_rows(worksheet)
                    if rows:
                        options.append(DataOption(sheet_name, rows))
                finally:
                    self._release_com_object(worksheet)
            self._release_com_object(worksheets)
        finally:
            if workbook is not None:
                try:
                    self._com_call(workbook, 'Close', False)
                except Exception:
                    pass
            if excel_app is not None:
                try:
                    self._com_call(excel_app, 'Quit')
                except Exception:
                    pass

            self._release_com_object(workbook)
            self._release_com_object(excel_app)

        return options

    def pick_input_file(self):
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

    def get_table_data(self, file_path):
        extension = os.path.splitext(file_path)[1].lower()

        if extension == '.csv':
            rows = self.read_delimited_file(file_path)
            return os.path.splitext(os.path.basename(file_path))[0], rows

        if extension not in ('.xlsx', '.xls'):
            raise Exception('Unsupported file type: {}'.format(extension))

        sheets = self.read_excel_workbook(file_path)
        if not sheets:
            raise Exception(
                'The workbook does not contain any populated worksheets.')

        selected = sheets[0]
        if len(sheets) > 1:
            selected = forms.SelectFromList.show(
                sheets,
                name_attr='display_name',
                multiselect=False,
                title='Select Worksheet'
            )
            if not selected:
                return None, None

        return selected.name, selected.rows

    @staticmethod
    def get_body_section(schedule):
        return schedule.GetTableData().GetSectionData(SectionType.Body)

    def collect_schedules(self):
        schedules = []
        for schedule in FilteredElementCollector(self.doc).OfClass(ViewSchedule):
            try:
                if schedule.IsTemplate:
                    continue
                if schedule.Name.startswith('<'):
                    continue
            except Exception:
                pass

            try:
                definition = schedule.Definition
                detail = 'Key Schedule' if definition.IsKeySchedule else 'Schedule'
                schedules.append(ScheduleOption(schedule, detail))
            except Exception:
                pass
        return sorted(schedules, key=lambda option: option.display_name.lower())

    def pick_schedules(self):
        if isinstance(self.active_view, ViewSchedule):
            return [self.active_view]

        schedules = self.collect_schedules()
        if not schedules:
            return []

        selected = forms.SelectFromList.show(
            schedules,
            name_attr='display_name',
            multiselect=True,
            title='Select Schedules to Populate'
        )
        if not selected:
            return []
        return [opt.schedule for opt in selected]

    def get_candidate_source_keys(self, schedule_header_key):
        candidates = [schedule_header_key]
        alias_values = self.source_header_aliases.get(schedule_header_key, [])
        for alias in alias_values:
            alias_key = self.normalize_header(alias)
            if alias_key and alias_key not in candidates:
                candidates.append(alias_key)
        return candidates

    def get_schedule_headers(self, schedule):
        definition = schedule.Definition
        headers = []
        for field_id in definition.GetFieldOrder():
            field = definition.GetField(field_id)
            headers.append(self.clean_cell_value(field.ColumnHeading))
        return headers

    def build_column_map(self, schedule_headers, source_headers):
        source_lookup = {}
        for index, header in enumerate(source_headers):
            key = self.normalize_header(header)
            if key and key not in source_lookup:
                source_lookup[key] = index

        mapped_columns = []
        missing_headers = []
        for schedule_col_index, schedule_header in enumerate(schedule_headers):
            schedule_key = self.normalize_header(schedule_header)
            source_col_index = None
            for candidate_key in self.get_candidate_source_keys(schedule_key):
                if candidate_key in source_lookup:
                    source_col_index = source_lookup[candidate_key]
                    break

            if source_col_index is not None:
                mapped_columns.append(
                    (schedule_col_index, source_col_index, schedule_header))
            else:
                missing_headers.append(schedule_header)

        return mapped_columns, missing_headers

    def get_source_header_index(self, source_headers, expected_name):
        normalized_headers = [self.normalize_header(h) for h in source_headers]
        normalized_expected = self.normalize_header(expected_name)
        if normalized_expected in normalized_headers:
            return normalized_headers.index(normalized_expected)
        return None

    @staticmethod
    def get_schedule_field_map(schedule):
        definition = schedule.Definition
        field_map = {}
        for col_index, field_id in enumerate(definition.GetFieldOrder()):
            field_map[col_index] = definition.GetField(field_id)
        return field_map

    def get_field_lookup_names(self, field):
        names = []
        try:
            internal_name = self.clean_cell_value(field.GetName())
            if internal_name:
                names.append(internal_name)
        except Exception:
            pass

        heading_name = self.clean_cell_value(field.ColumnHeading)
        if heading_name and heading_name not in names:
            names.append(heading_name)

        return names

    def get_param_target_and_param(self, element, param_id, fallback_names):
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

        element_type = self.doc.GetElement(type_id)
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

    def build_element_label_map(self, schedule, label_field):
        label_param_id = label_field.ParameterId
        fallback_names = self.get_field_lookup_names(label_field)
        element_map = {}
        total_elements = 0
        elements_with_labels = 0

        # type: ignore
        # type: ignore
        # type: ignore
        for element in FilteredElementCollector(self.doc, schedule.Id).WhereElementIsNotElementType():
            total_elements += 1
            try:
                _, param = self.get_param_target_and_param(
                    element, label_param_id, fallback_names)
                if param is not None:
                    label = self.clean_cell_value(
                        param.AsString() or param.AsValueString())
                    if label:
                        elements_with_labels += 1
                        key = label.lower()
                        if key not in element_map:
                            element_map[key] = element
            except Exception:
                pass

        element_map['_debug_total_elements'] = total_elements
        element_map['_debug_elements_with_labels'] = elements_with_labels
        return element_map

    @staticmethod
    def parse_number(value_text):
        cleaned = value_text.replace(',', '').strip()
        i = len(cleaned)
        while i > 0 and not (cleaned[i - 1].isdigit() or cleaned[i - 1] == '.'):
            i -= 1
        cleaned = cleaned[:i].strip()
        return float(cleaned)

    @staticmethod
    def _safe_float(text):
        try:
            return float(text)
        except Exception:
            return None

    def parse_length_inches(self, value_text):
        text = self.clean_cell_value(value_text)
        if not text:
            return 0.0

        normalized = text.replace(' ', '').replace('"', '').replace(
            '”', '').replace('′', "'").replace('’', "'")

        if "'" in normalized:
            feet_part, inches_part = normalized.split("'", 1)
            feet = self._safe_float(feet_part)
            if feet is None:
                return self.parse_number(text)

            inches_part = inches_part.strip()
            if inches_part.startswith('-'):
                inches_part = inches_part[1:]
            inches = self._safe_float(inches_part) if inches_part else 0.0
            if inches is None:
                inches = 0.0
            return (feet * 12.0) + inches

        if '-' in normalized:
            parts = normalized.split('-')
            if len(parts) == 2:
                feet = self._safe_float(parts[0])
                inches = self._safe_float(parts[1])
                if feet is not None and inches is not None:
                    return (feet * 12.0) + inches

        return self.parse_number(text)

    @staticmethod
    def is_length_parameter(param):
        try:
            definition = param.Definition
            if definition is None:
                return False

            try:
                data_type = definition.GetDataType()
                if SpecTypeId is not None and data_type == SpecTypeId.Length:
                    return True
            except Exception:
                pass

            try:
                return str(definition.ParameterType) == 'Length'
            except Exception:
                return False
        except Exception:
            return False

    def to_internal_double_value(self, param, numeric_value):
        if not self.is_length_parameter(param):
            return numeric_value

        if UnitUtils is not None and UnitTypeId is not None:
            try:
                return UnitUtils.ConvertToInternalUnits(numeric_value, UnitTypeId.Inches)
            except Exception:
                pass

        return numeric_value / 12.0

    @staticmethod
    def clear_parameter_value(param):
        """Try to clear a parameter value (null-like behavior in Revit)."""
        try:
            clear_value = getattr(param, 'ClearValue', None)
            if callable(clear_value):
                return bool(clear_value())
        except Exception:
            pass

        try:
            return bool(param.SetValueString(''))
        except Exception:
            pass

        return False

    def set_element_parameter(self, element, field, value_text):
        _, param = self.get_param_target_and_param(
            element, field.ParameterId, self.get_field_lookup_names(field))
        if param is None or param.IsReadOnly:
            return False
        try:
            storage = param.StorageType
            if storage == StorageType.String:
                param.Set(value_text)
            elif storage == StorageType.Integer:
                if value_text:
                    param.Set(int(self.parse_number(value_text)))
                else:
                    return self.clear_parameter_value(param)
            elif storage == StorageType.Double:
                if value_text:
                    if self.is_length_parameter(param):
                        numeric_value = self.parse_length_inches(value_text)
                    else:
                        numeric_value = self.parse_number(value_text)
                    param.Set(self.to_internal_double_value(
                        param, numeric_value))
                else:
                    return self.clear_parameter_value(param)
            else:
                return False
            return True
        except Exception:
            return False

    def update_schedule_rows_by_label(self, schedule, data_rows, mapped_columns, label_schedule_col, label_source_col):
        field_map = self.get_schedule_field_map(schedule)
        label_field = field_map.get(label_schedule_col)
        element_map = self.build_element_label_map(schedule, label_field)

        updated = 0
        empty_label = 0
        unmatched_source_labels = []
        matched_details = []
        matched_keys = set()

        for row_values in data_rows:
            label = row_values[label_source_col] if label_source_col < len(
                row_values) else ''
            label = self.clean_cell_value(label)
            if not label:
                empty_label += 1
                continue
            element = element_map.get(label.lower())
            if element is None:
                unmatched_source_labels.append(label)
                continue
            empty_cols = []
            for schedule_col_index, source_col_index, col_header in mapped_columns:
                if schedule_col_index == label_schedule_col:
                    continue
                field = field_map.get(schedule_col_index)
                if field is None:
                    continue
                cell_text = row_values[source_col_index] if source_col_index < len(
                    row_values) else ''
                if not cell_text:
                    empty_cols.append(col_header)
                self.set_element_parameter(element, field, cell_text)
            matched_keys.add(label.lower())
            matched_details.append((label, empty_cols))
            updated += 1

        unmatched_schedule_labels = sorted(
            k for k in element_map.keys() if k not in matched_keys
        )

        cleared_missing_labels = 0
        for missing_key in unmatched_schedule_labels:
            element = element_map.get(missing_key)
            if element is None:
                continue
            for schedule_col_index, source_col_index, col_header in mapped_columns:
                if schedule_col_index == label_schedule_col:
                    continue
                field = field_map.get(schedule_col_index)
                if field is None:
                    continue
                self.set_element_parameter(element, field, '')
            cleared_missing_labels += 1

        return (
            updated,
            empty_label,
            unmatched_source_labels,
            element_map,
            matched_details,
            unmatched_schedule_labels,
            cleared_missing_labels,
        )

    def get_import_rows(self, table_rows):
        if len(table_rows) < 2:
            raise Exception(
                'The source data must include a header row and at least one data row.')

        headers = [self.clean_cell_value(value) for value in table_rows[0]]
        data_rows = []
        for row in table_rows[1:]:
            normalized = [self.clean_cell_value(value) for value in row]
            if any(normalized):
                data_rows.append(normalized)

        if not data_rows:
            raise Exception(
                'No populated data rows were found below the header row.')

        return headers, data_rows

    def populate_schedule(self, schedule, source_headers, data_rows):
        schedule_headers = self.get_schedule_headers(schedule)
        if not schedule_headers:
            self.output.print_md(
                'Could not read schedule column headers - skipped.')
            return

        mapped_columns, missing_headers = self.build_column_map(
            schedule_headers, source_headers)
        if not mapped_columns:
            self.output.print_md(
                '**No source columns match the schedule headers - skipped.**')
            self.output.print_md(
                '- Schedule headers: {}'.format(', '.join(['`{}`'.format(h) for h in schedule_headers])))
            self.output.print_md(
                '- Excel headers: {}'.format(', '.join(['`{}`'.format(h) for h in source_headers])))
            return

        label_schedule_col = None
        for sched_col, src_col, header_name in mapped_columns:
            if self.normalize_header(header_name) == 'label':
                label_schedule_col = sched_col
                break

        label_source_col = self.get_source_header_index(
            source_headers, 'Label')

        if label_schedule_col is None or label_source_col is None:
            self.output.print_md(
                'No Label column found in both schedule and source - skipped.')
            self.output.print_md(
                '- Mapped columns: {}'.format(', '.join([item[2] for item in mapped_columns])))
            self.output.print_md(
                '- Source headers: {}'.format(', '.join(source_headers)))
            return

        with revit.Transaction('Populate Schedule: {}'.format(schedule.Name)):
            updated_count, empty_label_count, unmatched_source_labels, schedule_label_map, matched_details, unmatched_schedule_labels, cleared_missing_labels = self.update_schedule_rows_by_label(
                schedule, data_rows, mapped_columns, label_schedule_col, label_source_col
            )
            schedule.RefreshData()

        self.output.print_md(
            '- Schedule: {}'.format(self.output.linkify(schedule.Id)))
        self.output.print_md(
            '- Mapped columns: {}'.format(', '.join([item[2] for item in mapped_columns])))

        debug_total = schedule_label_map.pop('_debug_total_elements', 0)
        debug_with_labels = schedule_label_map.pop(
            '_debug_elements_with_labels', 0)
        self.output.print_md(
            '- Schedule elements collected: {} (with readable Label: {})'.format(debug_total, debug_with_labels))

        if missing_headers:
            self.output.print_md(
                '- Schedule columns not in source: {}'.format(', '.join(missing_headers)))
        self.output.print_md(
            '- Rows updated: {} / {}'.format(updated_count, len(schedule_label_map)))
        self.output.print_md(
            '- Missing schedule labels cleared: {}'.format(cleared_missing_labels))
        self.output.print_md(
            '- Rows with empty Label in source: {}'.format(empty_label_count))

        if matched_details:
            data_cols = [
                h for _, _, h in mapped_columns if self.normalize_header(h) != 'label']
            header_row = '| Label | ' + ' | '.join(data_cols) + ' |'
            sep_row = '|---|' + '|'.join(['---'] * len(data_cols)) + '|'
            self.output.print_md(header_row)
            self.output.print_md(sep_row)
            for lbl, empty_cols in sorted(matched_details, key=lambda x: x[0].lower()):
                cells = []
                for col in data_cols:
                    cells.append('*(empty)*' if col in empty_cols else 'OK')
                self.output.print_md('| {} | '.format(
                    lbl) + ' | '.join(cells) + ' |')

        if unmatched_schedule_labels:
            self.output.print_md('\n**Schedule labels with no CSV match ({}):** {}'.format(
                len(unmatched_schedule_labels),
                ', '.join(unmatched_schedule_labels)
            ))

        if unmatched_source_labels:
            self.output.print_md('\n**Source labels not in schedule ({}):** {}'.format(
                len(unmatched_source_labels),
                ', '.join(sorted(set(unmatched_source_labels)))
            ))

    def run_populate_from_file_dialog(self):
        file_path = self.pick_input_file()
        if not file_path:
            return

        worksheet_name, table_rows = self.get_table_data(file_path)
        if not table_rows:
            self.output.print_md('No data found in the selected file.')
            return

        selected_schedules = self.pick_schedules()
        if not selected_schedules:
            return

        source_headers, data_rows = self.get_import_rows(table_rows)

        self.output.print_md('# Populate Schedule From Excel')
        self.output.print_md('- Source: {}'.format(file_path))
        self.output.print_md('- Worksheet: {}'.format(worksheet_name))
        self.output.print_md(
            '- Schedules selected: {}'.format(len(selected_schedules)))

        for schedule in selected_schedules:
            self.output.print_md('---')
            self.output.print_md('## {}'.format(schedule.Name))
            self.populate_schedule(schedule, source_headers, data_rows)


class RevitSchedule(RevitSchedules):
    """Backwards-compatible alias."""
