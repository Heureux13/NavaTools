# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import revit, script
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    FilteredElementCollector,
    View,
    ViewDuplicateOption,
)
from System.Windows.Forms import Button, DialogResult, Form, Label, TextBox, TreeNode, TreeView
import clr

clr.AddReference("System.Windows.Forms")


# Button info
# ======================================================================
__title__ = 'Create Views'
__doc__ = '''
Create views from selected scope boxes
'''


# Variables
# ======================================================================
doc = revit.doc
output = script.get_output()


def get_element_id_value(element_id):
    try:
        return element_id.Value
    except Exception:
        return element_id.IntegerValue


class MultiPickerForm(Form):
    def __init__(self, title, action_text, elements, get_name, get_id, prechecked_ids=None):
        Form.__init__(self)
        self.Text = title
        self.Width = 560
        self.Height = 670

        self.elements = sorted(elements, key=lambda element: get_name(element).lower())
        self.get_name = get_name
        self.get_id = get_id
        self.checked_ids = set(prechecked_ids or [])

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
        btn_ok.Text = action_text
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
        checked_ids = self.checked_ids

        for element in self.elements:
            name = self.get_name(element)
            if filter_text and filter_text not in name.lower():
                continue

            element_id_value = self.get_id(element)
            node = TreeNode(name)
            node.Tag = element_id_value
            node.Checked = element_id_value in checked_ids
            self.tree_view.Nodes.Add(node)

    def _on_filter_changed(self, sender, args):
        filter_text = (sender.Text or "").strip().lower()
        self._build_tree(filter_text if filter_text else None)

    def _on_node_checked(self, sender, args):
        element_id = args.Node.Tag
        if element_id is None:
            return

        if args.Node.Checked:
            self.checked_ids.add(element_id)
        else:
            self.checked_ids.discard(element_id)

    def _on_select_all(self, sender, args):
        for element in self.elements:
            self.checked_ids.add(self.get_id(element))
        self._build_tree((self.search_box.Text or "").strip().lower() or None)

    def _on_clear(self, sender, args):
        self.checked_ids.clear()
        self._build_tree((self.search_box.Text or "").strip().lower() or None)

    def get_checked_ids(self):
        return list(self.checked_ids)


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


def collect_selected_source_views():
    selected_views = []

    try:
        for element in revit.get_selection():
            if isinstance(element, View) and not element.IsTemplate and can_duplicate_as_dependent(element):
                selected_views.append(element)
    except Exception:
        pass

    return sorted(selected_views, key=lambda view: view.Name.lower())


def duplicate_view_as_dependent(view):
    if not view.CanViewBeDuplicated(ViewDuplicateOption.AsDependent):
        raise Exception("View cannot be duplicated as dependent.")
    new_view_id = view.Duplicate(ViewDuplicateOption.AsDependent)
    return doc.GetElement(new_view_id)


selected_source_views = collect_selected_source_views()
if not selected_source_views:
    TaskDialog.Show(
        "Create Views",
        "No valid views selected. In the Project Browser, select one or more non-template views that can be duplicated as dependent, then run the command again."
    )
    script.exit()

all_scope_boxes = collect_scope_boxes(doc)
if not all_scope_boxes:
    TaskDialog.Show("Create Views", "No scope boxes found in this project.")
    script.exit()

scope_form = MultiPickerForm(
    title="Select Scope Boxes",
    action_text="Create Views",
    elements=all_scope_boxes,
    get_name=lambda scope_box: scope_box.Name,
    get_id=lambda scope_box: get_element_id_value(scope_box.Id),
)

result = scope_form.ShowDialog()
if result != DialogResult.OK:
    script.exit()

selected_id_values = set(scope_form.get_checked_ids())
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
    for source_view in selected_source_views:
        for scope_box in selected_scope_boxes:
            try:
                dup_view = duplicate_view_as_dependent(source_view)

                scope_param = dup_view.get_Parameter(
                    BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)
                if not scope_param or scope_param.IsReadOnly:
                    doc.Delete(dup_view.Id)
                    failed.append(
                        "{} - {} (cannot set scope box)".format(source_view.Name, scope_box.Name))
                    continue

                scope_param.Set(scope_box.Id)

                proposed_name = "{} - {}".format(source_view.Name, scope_box.Name)
                unique_name = get_unique_view_name(proposed_name, existing_names)
                dup_view.Name = unique_name
                created.append((dup_view, source_view.Name, scope_box.Name))
            except Exception as err:
                failed.append("{} - {} ({})".format(source_view.Name, scope_box.Name, err))

if created:
    output.print_md("# Created {} view(s)".format(len(created)))
    for new_view, source_name, scope_name in created:
        output.print_md("- {} - {} -> {}".format(source_name, scope_name,
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
