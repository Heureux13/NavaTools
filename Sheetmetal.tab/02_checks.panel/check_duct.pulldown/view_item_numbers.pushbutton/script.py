# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import revit, script
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    ElementId,
    VisibleInViewFilter,
)
from Autodesk.Revit.UI import TaskDialog
from System.Collections.Generic import List
from System.Windows.Forms import (
    Form,
    Label,
    Button,
    DialogResult,
    FormStartPosition,
    TextBox,
    TreeView,
    TreeNode,
)
import clr
import re

clr.AddReference("System.Windows.Forms")


# Button info
# ===================================================
__title__ = "View Item Numbers"
__doc__ = """
View Item Number values in the active view using a tree grouped by 100 ranges.
"""


# Revit context
# ===================================================
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
view = revit.active_view
output = script.get_output()


def natural_sort_key(value):
    text = "" if value is None else str(value)
    return [int(part) if part.isdigit() else part.lower() for part in re.split(r"(\d+)", text)]


def get_param_text(param):
    if not param:
        return ""
    try:
        val = param.AsString()
        if not val:
            val = param.AsValueString()
        return val.strip() if val else ""
    except Exception:
        return ""


def get_param_value_by_names(elem, param_names):
    for name in param_names:
        try:
            p = elem.LookupParameter(name)
            txt = get_param_text(p)
            if txt:
                return txt
        except Exception:
            continue
    return ""


def parse_item_number(value_text):
    text = (value_text or "").strip()
    if not text:
        return None
    try:
        num = float(text)
    except Exception:
        return None
    if abs(num - int(round(num))) > 1e-9:
        return None
    return int(round(num))


def get_bucket_bounds(item_number):
    # First bucket is 0-100, then 101-200, 201-300, etc.
    if item_number <= 100:
        return 0, 100
    start = ((item_number - 1) // 100) * 100 + 1
    end = start + 99
    return start, end


def collect_fab_ducts_in_view(doc_obj, view_obj):
    return list(
        FilteredElementCollector(doc_obj, view_obj.Id)
        .OfCategory(BuiltInCategory.OST_FabricationDuctwork)
        .WhereElementIsNotElementType()
        .WherePasses(VisibleInViewFilter(doc_obj, view_obj.Id))
        .ToElements()
    )


def build_item_number_groups(elements):
    """
    Build structure:
      {
            (start, end): {
              item_number_int: [elements]
            }
      }
    """
    buckets = {}

    for elem in elements:
        try:
            p = elem.LookupParameter("Item Number")
            raw_value = get_param_text(p)
            item_number = parse_item_number(raw_value)
            if item_number is None:
                continue
        except Exception:
            continue

        bucket_key = get_bucket_bounds(item_number)
        if bucket_key not in buckets:
            buckets[bucket_key] = {}
        if item_number not in buckets[bucket_key]:
            buckets[bucket_key][item_number] = []
        buckets[bucket_key][item_number].append(elem)

    return buckets


class ItemNumberTreePicker(Form):
    def __init__(self, buckets):
        Form.__init__(self)
        self.Text = "View Item Numbers"
        self.StartPosition = FormStartPosition.CenterScreen
        self.Width = 760
        self.Height = 760
        self.buckets = buckets

        self.lbl_prompt = Label()
        self.lbl_prompt.Text = "Select one or more Item Number values from grouped ranges."
        self.lbl_prompt.Top = 10
        self.lbl_prompt.Left = 20
        self.lbl_prompt.Width = 700
        self.lbl_prompt.Height = 36
        self.lbl_prompt.AutoSize = False
        self.Controls.Add(self.lbl_prompt)

        self.search_box = TextBox()
        self.search_box.Top = 50
        self.search_box.Left = 20
        self.search_box.Width = 700
        self.search_box.TextChanged += self._on_search_changed
        self.Controls.Add(self.search_box)

        self.tree_view = TreeView()
        self.tree_view.Top = 80
        self.tree_view.Left = 20
        self.tree_view.Width = 700
        self.tree_view.Height = 590
        self.tree_view.HideSelection = False
        self.tree_view.CheckBoxes = True
        self.tree_view.AfterCheck += self._on_after_check
        self.Controls.Add(self.tree_view)

        self.lbl_selected = Label()
        self.lbl_selected.Text = "Selected: 0 item numbers"
        self.lbl_selected.Top = 680
        self.lbl_selected.Left = 20
        self.lbl_selected.Width = 500
        self.Controls.Add(self.lbl_selected)

        btn_ok = Button()
        btn_ok.Text = "Select"
        btn_ok.Top = 675
        btn_ok.Left = 560
        btn_ok.Width = 75
        btn_ok.DialogResult = DialogResult.OK
        btn_ok.Click += self._on_ok_clicked
        self.Controls.Add(btn_ok)
        self.AcceptButton = btn_ok

        btn_cancel = Button()
        btn_cancel.Text = "Cancel"
        btn_cancel.Top = 675
        btn_cancel.Left = 645
        btn_cancel.Width = 75
        btn_cancel.DialogResult = DialogResult.Cancel
        self.Controls.Add(btn_cancel)
        self.CancelButton = btn_cancel

        self._build_tree(None)

    def _build_tree(self, search_filter):
        self.tree_view.Nodes.Clear()
        search = (search_filter or "").strip().lower()

        bucket_keys = sorted(self.buckets.keys(), key=lambda b: (b[0], b[1]))
        for bucket_start, bucket_end in bucket_keys:
            numbers_map = self.buckets[(bucket_start, bucket_end)]

            bucket_label = "{}-{}".format(bucket_start, bucket_end)
            bucket_total = sum(len(numbers_map[n]) for n in numbers_map.keys())

            bucket_node = TreeNode("{} ({} elements)".format(bucket_label, bucket_total))
            bucket_node.Tag = ("bucket", bucket_start, bucket_end)

            for item_number in sorted(numbers_map.keys()):
                elems = numbers_map[item_number]
                item_text = str(item_number)
                haystack = "{} {}".format(bucket_label, item_text).lower()
                if search and search not in haystack:
                    continue

                item_node = TreeNode("{} ({} elements)".format(item_text, len(elems)))
                item_node.Tag = ("item", item_number)
                bucket_node.Nodes.Add(item_node)

            if bucket_node.Nodes.Count > 0:
                self.tree_view.Nodes.Add(bucket_node)

        self.tree_view.CollapseAll()

    def _on_search_changed(self, sender, args):
        self._build_tree(sender.Text)

    def _on_after_check(self, sender, args):
        node = args.Node
        if node is None:
            return

        self.tree_view.AfterCheck -= self._on_after_check
        try:
            if node.Tag and node.Tag[0] == "bucket":
                for child in node.Nodes:
                    child.Checked = node.Checked
        finally:
            self.tree_view.AfterCheck += self._on_after_check

        count_selected = len(self.get_checked_item_numbers())
        self.lbl_selected.Text = "Selected: {} item numbers".format(count_selected)

    def get_checked_item_numbers(self):
        checked = []
        for bucket_node in self.tree_view.Nodes:
            for item_node in bucket_node.Nodes:
                if item_node.Checked and item_node.Tag and item_node.Tag[0] == "item":
                    checked.append(int(item_node.Tag[1]))
        return checked

    def _on_ok_clicked(self, sender, args):
        if not self.get_checked_item_numbers():
            TaskDialog.Show("Select Item Numbers", "Check at least one Item Number before clicking Select.")
            self.DialogResult = getattr(DialogResult, "None")


def find_elements_for_item_numbers(buckets, selected_numbers):
    picked = []
    wanted = set(selected_numbers)
    for numbers_map in buckets.values():
        for number_value, elems in numbers_map.items():
            if number_value in wanted:
                picked.extend(elems)
    return picked


def get_item_number_for_elem(elem):
    txt = get_param_value_by_names(elem, ["Item Number"])
    parsed = parse_item_number(txt)
    return parsed if parsed is not None else -1


try:
    ducts = collect_fab_ducts_in_view(doc, view)
    if not ducts:
        TaskDialog.Show("No Ducts", "No fabrication ductwork found in this view.")
        script.exit()

    buckets = build_item_number_groups(ducts)
    if not buckets:
        TaskDialog.Show("No Item Numbers", "No numeric Item Number values found in this view.")
        script.exit()

    form = ItemNumberTreePicker(buckets)
    if form.ShowDialog() != DialogResult.OK:
        script.exit()

    selected_numbers = form.get_checked_item_numbers()
    if not selected_numbers:
        script.exit()

    selected_elems = find_elements_for_item_numbers(buckets, selected_numbers)
    if not selected_elems:
        TaskDialog.Show("No Matches", "No elements matched the selected Item Number values.")
        script.exit()

    selected_elems = sorted(
        selected_elems,
        key=lambda e: (get_item_number_for_elem(e), e.Id.IntegerValue),
    )

    selected_ids = List[ElementId]()
    unique_selected_elems = []
    seen_ids = set()
    for elem in selected_elems:
        try:
            eid_int = elem.Id.IntegerValue
            if eid_int in seen_ids:
                continue
            seen_ids.add(eid_int)
            selected_ids.Add(elem.Id)
            unique_selected_elems.append(elem)
        except Exception:
            continue

    if selected_ids.Count == 0:
        TaskDialog.Show("No Matches", "No valid element IDs were found from selected Item Numbers.")
        script.exit()

    uidoc.Selection.SetElementIds(selected_ids)

    selected_unique_numbers = sorted(set(selected_numbers))

    output.print_md("# Item Number Selection")
    output.print_md("- Selected item numbers: {}".format(len(selected_numbers)))
    output.print_md("- Selected elements: {}".format(selected_ids.Count))
    output.print_md("- Min item number: {}".format(selected_unique_numbers[0]))
    output.print_md("- Max item number: {}".format(selected_unique_numbers[-1]))

    output.print_md("## Indexed Selection")
    for idx, elem in enumerate(unique_selected_elems, start=1):
        index_text = "{:06d}".format(idx)
        item_number_value = get_param_value_by_names(elem, ["Item Number"]) or "(missing)"
        output.print_md("{} | {} | Item Number: {}".format(index_text, output.linkify(elem.Id), item_number_value))

except Exception as exc:
    TaskDialog.Show("Error", "Script failed: {}".format(str(exc)))
    import traceback
    output.print_md("```\n{}\n```".format(traceback.format_exc()))
