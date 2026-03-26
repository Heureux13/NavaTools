# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

import clr

clr.AddReference("System.Windows.Forms")

from System.Windows.Forms import Button, DialogResult, Form, Label, TextBox, TreeNode, TreeView
from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    FilteredElementCollector,
    View,
    ViewDuplicateOption,
)
from Autodesk.Revit.UI import TaskDialog
from pyrevit import revit, script


# Button info
# ======================================================================
__title__ = 'Create Views'
__doc__ = '''
Create views from selected scope boxes
'''


# Variables
# ======================================================================
doc = revit.doc
source_view = revit.active_view
output = script.get_output()


def get_element_id_value(element_id):
    try:
        return element_id.Value
    except Exception:
        return element_id.IntegerValue


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
        btn_ok.Text = "Create Views"
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


def can_duplicate_as_dependent(view):
    try:
        return view and view.CanViewBeDuplicated(ViewDuplicateOption.AsDependent)
    except Exception:
        return False


def get_unique_view_name(base_name, existing_names):
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


def collect_scope_boxes(document):
    return list(
        FilteredElementCollector(document)
        .OfCategory(BuiltInCategory.OST_VolumeOfInterest)
        .WhereElementIsNotElementType()
        .ToElements()
    )


def duplicate_view_as_dependent(view):
    if not view.CanViewBeDuplicated(ViewDuplicateOption.AsDependent):
        raise Exception("View cannot be duplicated as dependent.")
    new_view_id = view.Duplicate(ViewDuplicateOption.AsDependent)
    return doc.GetElement(new_view_id)


if not source_view:
    TaskDialog.Show("Create Views", "No active view found.")
    script.exit()

if not isinstance(source_view, View) or source_view.IsTemplate:
    TaskDialog.Show(
        "Create Views", "Active view is not a valid non-template view.")
    script.exit()

if not can_duplicate_as_dependent(source_view):
    TaskDialog.Show(
        "Create Views",
        "The active view cannot be duplicated as dependent."
    )
    script.exit()

all_scope_boxes = collect_scope_boxes(doc)
if not all_scope_boxes:
    TaskDialog.Show("Create Views", "No scope boxes found in this project.")
    script.exit()

form = ScopeBoxPickerForm(all_scope_boxes)
result = form.ShowDialog()
if result != DialogResult.OK:
    script.exit()

selected_id_values = set(form.get_checked_scope_ids())
if not selected_id_values:
    TaskDialog.Show("Create Views", "No scope boxes were selected.")
    script.exit()

selected_scope_boxes = [
    scope_box for scope_box in sorted(all_scope_boxes, key=lambda scope_box: scope_box.Name.lower())
    if get_element_id_value(scope_box.Id) in selected_id_values
]

existing_names = {
    view.Name for view in FilteredElementCollector(doc).OfClass(View).ToElements()
    if getattr(view, "Name", None)
}

created = []
failed = []

with revit.Transaction("Create Scope Box Views"):
    for scope_box in selected_scope_boxes:
        try:
            dup_view = duplicate_view_as_dependent(source_view)

            scope_param = dup_view.get_Parameter(
                BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)
            if not scope_param or scope_param.IsReadOnly:
                doc.Delete(dup_view.Id)
                failed.append(
                    "{} (cannot set scope box)".format(scope_box.Name))
                continue

            scope_param.Set(scope_box.Id)

            proposed_name = "{} - {}".format(source_view.Name, scope_box.Name)
            unique_name = get_unique_view_name(proposed_name, existing_names)
            dup_view.Name = unique_name
            created.append((dup_view, scope_box.Name))
        except Exception as err:
            failed.append("{} ({})".format(scope_box.Name, err))

if created:
    output.print_md("# Created {} view(s)".format(len(created)))
    for new_view, scope_name in created:
        output.print_md("- {} -> {}".format(scope_name,
                        output.linkify(new_view.Id)))

if failed:
    output.print_md("---")
    output.print_md("## Could not create {} view(s)".format(len(failed)))
    for message in failed:
        output.print_md("- {}".format(message))

TaskDialog.Show(
    "Create Views",
    "Created {} view(s).{}".format(
        len(created),
        " Failed: {}.".format(len(failed)) if failed else ""
    )
)
