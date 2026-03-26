# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import forms, revit, script
from Autodesk.Revit.DB import (
	BuiltInCategory,
	BuiltInParameter,
	FamilySymbol,
	FilteredElementCollector,
	View,
	ViewDuplicateOption,
	ViewPlan,
	Viewport,
	ViewSheet,
	XYZ,
)
from Autodesk.Revit.UI import TaskDialog
from System.Windows.Forms import Button, DialogResult, Form, Label, TextBox, TreeNode, TreeView
import clr

clr.AddReference("System.Windows.Forms")


# Button info
# ======================================================================
__title__ = 'Create Sheet'
__doc__ = '''
Create sheets from selected plan views and scope boxes.
'''


# Variables
# ======================================================================
doc = revit.doc
uidoc = __revit__.ActiveUIDocument
output = script.get_output()


class ListOption(object):
	def __init__(self, item, display_name):
		self.item = item
		self.display_name = display_name


def get_element_id_value(element_id):
	try:
		return element_id.Value
	except Exception:
		return element_id.IntegerValue


def get_level_name(view):
	level = getattr(view, "GenLevel", None)
	return level.Name if level else "No Level"


def get_view_display_name(view):
	return "{} | {}".format(get_level_name(view), view.Name)


def get_titleblock_display_name(symbol):
	family = getattr(symbol, "Family", None)
	family_name = family.Name if family else "Title Block"
	return "{} | {}".format(family_name, symbol.Name)


def get_unique_name(base_name, existing_names):
	if base_name not in existing_names:
		existing_names.add(base_name)
		return base_name

	index = 2
	while True:
		candidate = "{} ({})".format(base_name, index)
		if candidate not in existing_names:
			existing_names.add(candidate)
			return candidate
		index += 1


def get_unique_sheet_number(prefix, counter, existing_numbers):
	current = counter
	while True:
		candidate = "{}{:03d}".format(prefix, current)
		if candidate not in existing_numbers:
			existing_numbers.add(candidate)
			return candidate, current + 1
		current += 1


def collect_scope_boxes(document):
	return sorted(
		FilteredElementCollector(document)
		.OfCategory(BuiltInCategory.OST_VolumeOfInterest)
		.WhereElementIsNotElementType()
		.ToElements(),
		key=lambda scope_box: scope_box.Name.lower()
	)


def collect_titleblocks(document):
	return sorted(
		FilteredElementCollector(document)
		.OfClass(FamilySymbol)
		.OfCategory(BuiltInCategory.OST_TitleBlocks)
		.ToElements(),
		key=lambda symbol: get_titleblock_display_name(symbol).lower()
	)


def collect_selected_plan_views(document, active_uidoc):
	selected = []
	for element_id in active_uidoc.Selection.GetElementIds():
		element = document.GetElement(element_id)
		if isinstance(element, ViewPlan) and not element.IsTemplate:
			selected.append(element)
	return sorted(selected, key=lambda view: get_view_display_name(view).lower())


def collect_all_plan_views(document):
	return sorted(
		[
			view for view in FilteredElementCollector(document).OfClass(ViewPlan)
			if not view.IsTemplate
		],
		key=lambda view: get_view_display_name(view).lower()
	)


def duplicate_view_as_dependent(view):
	if not view.CanViewBeDuplicated(ViewDuplicateOption.AsDependent):
		raise Exception("View cannot be duplicated as dependent.")
	new_view_id = view.Duplicate(ViewDuplicateOption.AsDependent)
	return doc.GetElement(new_view_id)


def get_sheet_center(sheet):
	outline = sheet.Outline
	return XYZ(
		(outline.Min.U + outline.Max.U) / 2.0,
		(outline.Min.V + outline.Max.V) / 2.0,
		0
	)


class ScopeBoxPickerForm(Form):
	def __init__(self, scope_boxes):
		Form.__init__(self)
		self.Text = "Select Scope Boxes"
		self.Width = 560
		self.Height = 670

		self.scope_boxes = sorted(
			scope_boxes, key=lambda scope_box: scope_box.Name.lower())
		self.checked_scope_ids = set()

		label = Label()
		label.Text = "Search"
		label.Left = 20
		label.Top = 18
		label.Width = 80
		self.Controls.Add(label)

		self.search_box = TextBox()
		self.search_box.Left = 20
		self.search_box.Top = 38
		self.search_box.Width = 500
		self.search_box.TextChanged += self._on_filter_changed
		self.Controls.Add(self.search_box)

		self.tree_view = TreeView()
		self.tree_view.Left = 20
		self.tree_view.Top = 70
		self.tree_view.Width = 500
		self.tree_view.Height = 500
		self.tree_view.CheckBoxes = True
		self.tree_view.AfterCheck += self._on_node_checked
		self.Controls.Add(self.tree_view)

		btn_select_all = Button()
		btn_select_all.Text = "Select All"
		btn_select_all.Left = 20
		btn_select_all.Top = 585
		btn_select_all.Width = 115
		btn_select_all.Click += self._on_select_all
		self.Controls.Add(btn_select_all)

		btn_clear = Button()
		btn_clear.Text = "Clear"
		btn_clear.Left = 145
		btn_clear.Top = 585
		btn_clear.Width = 90
		btn_clear.Click += self._on_clear
		self.Controls.Add(btn_clear)

		btn_ok = Button()
		btn_ok.Text = "Create Sheets"
		btn_ok.Left = 320
		btn_ok.Top = 585
		btn_ok.Width = 95
		btn_ok.DialogResult = DialogResult.OK
		self.Controls.Add(btn_ok)
		self.AcceptButton = btn_ok

		btn_cancel = Button()
		btn_cancel.Text = "Cancel"
		btn_cancel.Left = 425
		btn_cancel.Top = 585
		btn_cancel.Width = 95
		btn_cancel.DialogResult = DialogResult.Cancel
		self.Controls.Add(btn_cancel)
		self.CancelButton = btn_cancel

		self._build_tree()

	def _build_tree(self, filter_text=None):
		self.tree_view.Nodes.Clear()
		checked_ids = self.checked_scope_ids

		for scope_box in self.scope_boxes:
			name = scope_box.Name
			if filter_text and filter_text not in name.lower():
				continue

			scope_id_value = get_element_id_value(scope_box.Id)
			node = TreeNode(name)
			node.Tag = scope_id_value
			node.Checked = scope_id_value in checked_ids
			self.tree_view.Nodes.Add(node)

	def _on_filter_changed(self, sender, args):
		filter_text = (sender.Text or "").strip().lower()
		self._build_tree(filter_text if filter_text else None)

	def _on_node_checked(self, sender, args):
		scope_id = args.Node.Tag
		if scope_id is None:
			return

		if args.Node.Checked:
			self.checked_scope_ids.add(scope_id)
		else:
			self.checked_scope_ids.discard(scope_id)

	def _on_select_all(self, sender, args):
		for scope_box in self.scope_boxes:
			self.checked_scope_ids.add(get_element_id_value(scope_box.Id))
		self._build_tree((self.search_box.Text or "").strip().lower() or None)

	def _on_clear(self, sender, args):
		self.checked_scope_ids.clear()
		self._build_tree((self.search_box.Text or "").strip().lower() or None)

	def get_checked_scope_ids(self):
		return list(self.checked_scope_ids)


selected_views = collect_selected_plan_views(doc, uidoc)
if not selected_views:
	all_views = collect_all_plan_views(doc)
	if not all_views:
		forms.alert("No plan views found in this project.", exitscript=True)

	selected_view_options = [
		ListOption(view, get_view_display_name(view)) for view in all_views
	]
	picked_view_options = forms.SelectFromList.show(
		selected_view_options,
		name_attr='display_name',
		multiselect=True,
		title='Select Plan Views'
	)
	if not picked_view_options:
		script.exit()
	selected_views = [option.item for option in picked_view_options]

all_scope_boxes = collect_scope_boxes(doc)
if not all_scope_boxes:
	forms.alert("No scope boxes found in this project.", exitscript=True)

scope_form = ScopeBoxPickerForm(all_scope_boxes)
scope_result = scope_form.ShowDialog()
if scope_result != DialogResult.OK:
	script.exit()

selected_scope_id_values = set(scope_form.get_checked_scope_ids())
if not selected_scope_id_values:
	forms.alert("No scope boxes selected.", exitscript=True)

selected_scope_boxes = [
	scope_box for scope_box in all_scope_boxes
	if get_element_id_value(scope_box.Id) in selected_scope_id_values
]

titleblocks = collect_titleblocks(doc)
if not titleblocks:
	forms.alert("No title blocks found in this project.", exitscript=True)

titleblock_options = [
	ListOption(symbol, get_titleblock_display_name(symbol)) for symbol in titleblocks
]
selected_titleblock = forms.SelectFromList.show(
	titleblock_options,
	name_attr='display_name',
	multiselect=False,
	title='Select Title Block'
)
if not selected_titleblock:
	script.exit()

sheet_prefix = forms.ask_for_string(
	default='M-',
	prompt='Enter a sheet number prefix. Sheet numbers will be generated as PREFIX001, PREFIX002, ...',
	title='Sheet Number Prefix'
)
if sheet_prefix is None:
	script.exit()

sheet_prefix = sheet_prefix.strip()
if not sheet_prefix:
	forms.alert("Sheet number prefix cannot be blank.", exitscript=True)

existing_view_names = {
	view.Name for view in FilteredElementCollector(doc).OfClass(View).ToElements()
	if getattr(view, 'Name', None)
}
existing_sheet_numbers = {
	sheet.SheetNumber for sheet in FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
	if getattr(sheet, 'SheetNumber', None)
}

created = []
failed = []
sheet_counter = 1

with revit.Transaction('Create Sheets From Scope Boxes'):
	for source_view in selected_views:
		if not source_view.CanViewBeDuplicated(ViewDuplicateOption.AsDependent):
			failed.append('{} (cannot duplicate as dependent)'.format(source_view.Name))
			continue

		for scope_box in selected_scope_boxes:
			dup_view = None
			sheet = None
			try:
				dup_view = duplicate_view_as_dependent(source_view)

				scope_param = dup_view.get_Parameter(
					BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)
				if not scope_param or scope_param.IsReadOnly:
					raise Exception('cannot set scope box')

				scope_param.Set(scope_box.Id)

				proposed_view_name = '{} - {}'.format(source_view.Name, scope_box.Name)
				dup_view.Name = get_unique_name(proposed_view_name, existing_view_names)

				sheet = ViewSheet.Create(doc, selected_titleblock.item.Id)
				sheet_number, sheet_counter = get_unique_sheet_number(
					sheet_prefix,
					sheet_counter,
					existing_sheet_numbers
				)
				sheet.SheetNumber = sheet_number
				sheet.Name = dup_view.Name

				if not Viewport.CanAddViewToSheet(doc, sheet.Id, dup_view.Id):
					raise Exception('view cannot be placed on sheet')

				Viewport.Create(doc, sheet.Id, dup_view.Id, get_sheet_center(sheet))

				created.append((sheet, dup_view, source_view.Name, scope_box.Name))
			except Exception as err:
				if sheet is not None:
					doc.Delete(sheet.Id)
				if dup_view is not None:
					doc.Delete(dup_view.Id)
				failed.append(
					'{} | {} ({})'.format(source_view.Name, scope_box.Name, err)
				)

if created:
	output.print_md('# Created {} sheet(s)'.format(len(created)))
	for sheet, new_view, source_name, scope_name in created:
		output.print_md(
			'- {} | {} -> {} | {}'.format(
				source_name,
				scope_name,
				output.linkify(sheet.Id),
				output.linkify(new_view.Id)
			)
		)

if failed:
	output.print_md('---')
	output.print_md('## Could not create {} sheet(s)'.format(len(failed)))
	for message in failed:
		output.print_md('- {}'.format(message))

TaskDialog.Show(
	'Create Sheet',
	'Created {} sheet(s).{}'.format(
		len(created),
		' Failed: {}.'.format(len(failed)) if failed else ''
	)
)
