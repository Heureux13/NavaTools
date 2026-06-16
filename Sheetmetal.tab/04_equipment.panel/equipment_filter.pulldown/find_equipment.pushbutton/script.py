# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from constants.print_outputs import print_disclaimer
from config.parameters_registry import BBM_LABEL
from pyrevit import revit, script
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, ElementId
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
__title__ = "Find Equipment by BBM Label"
__doc__ = """
Tree picker for Mechanical Equipment grouped by BBM label.
"""


# Revit context
# ===================================================
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
view = revit.active_view
output = script.get_output()


def format_index(num):
    return str(int(num)).zfill(4)


def natural_sort_key(value):
    text = "" if value is None else str(value)
    return [int(part) if part.isdigit() else part.lower() for part in re.split(r"(\d+)", text)]


def get_param_case_insensitive(element, param_name):
    target = (param_name or "").strip().lower()
    if not target or element is None:
        return None
    try:
        for param in element.Parameters:
            try:
                definition = param.Definition
                if definition and definition.Name and definition.Name.strip().lower() == target:
                    return param
            except Exception:
                continue
    except Exception:
        return None
    return None


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


def get_type_identity(instance, doc_obj):
    family_name = "(No Family Name)"
    type_name = "(No Type Name)"

    try:
        type_elem = doc_obj.GetElement(instance.GetTypeId())
    except Exception:
        type_elem = None

    if type_elem is None:
        return family_name, type_name

    try:
        fam = getattr(type_elem, "Family", None)
        if fam and fam.Name:
            family_name = fam.Name.strip()
    except Exception:
        pass

    if family_name == "(No Family Name)":
        try:
            fam_name = getattr(type_elem, "FamilyName", "")
            if fam_name:
                family_name = fam_name.strip()
        except Exception:
            pass

    try:
        tname = getattr(type_elem, "Name", "")
        if tname:
            type_name = tname.strip()
    except Exception:
        pass

    return family_name, type_name


def get_bbm_label_value(instance, doc_obj):
    names = ("_UMI_BBM_Lable", BBM_LABEL, "_UMI_BBM_Label")

    for pname in names:
        value = get_param_text(get_param_case_insensitive(instance, pname))
        if value:
            return value

    try:
        type_elem = doc_obj.GetElement(instance.GetTypeId())
    except Exception:
        type_elem = None

    if type_elem is not None:
        for pname in names:
            value = get_param_text(
                get_param_case_insensitive(type_elem, pname))
            if value:
                return value

    return "(empty)"


def collect_view_equipment_instances(doc_obj, view_obj):
    return list(
        FilteredElementCollector(doc_obj, view_obj.Id)
        .OfCategory(BuiltInCategory.OST_MechanicalEquipment)
        .WhereElementIsNotElementType()
        .ToElements()
    )


def get_label_group_key(label):
    text = (label or "").strip()
    if not text:
        return "(empty)"

    if "-" in text:
        prefix = text.split("-", 1)[0].strip()
        return prefix if prefix else text

    return text


def get_elevation_from_level(instance, doc_obj):
    """Get the 'Elevation from Level' parameter value."""
    return get_param_text(get_param_case_insensitive(instance, "Elevation from Level"))


def build_bbm_catalog(instances, doc_obj):
    labels_catalog = {}

    for inst in instances:
        if inst is None:
            continue

        label = get_bbm_label_value(inst, doc_obj)
        family_name, type_name = get_type_identity(inst, doc_obj)
        elem_id_int = inst.Id.IntegerValue
        elevation = get_elevation_from_level(inst, doc_obj)

        if label not in labels_catalog:
            labels_catalog[label] = []

        labels_catalog[label].append({
            "elem_id_int": elem_id_int,
            "family": family_name,
            "type": type_name,
            "elevation": elevation,
        })

    for label in labels_catalog.keys():
        labels_catalog[label] = sorted(
            labels_catalog[label],
            key=lambda r: natural_sort_key("{} {} {}".format(
                r["family"], r["type"], r["elem_id_int"])),
        )

    grouped_catalog = {}
    for label, rows in labels_catalog.items():
        group_key = get_label_group_key(label)
        if group_key not in grouped_catalog:
            grouped_catalog[group_key] = {}
        grouped_catalog[group_key][label] = rows

    return grouped_catalog


class EquipmentByLabelPicker(Form):
    def __init__(self, title, prompt, catalog):
        Form.__init__(self)
        self.Text = title
        self.StartPosition = FormStartPosition.CenterScreen
        self.Width = 840
        self.Height = 760
        self.catalog = catalog
        self.selected_element_ids = []
        self.selected_items = []

        self.lbl_prompt = Label()
        self.lbl_prompt.Text = prompt
        self.lbl_prompt.Top = 10
        self.lbl_prompt.Left = 20
        self.lbl_prompt.Width = 780
        self.lbl_prompt.Height = 42
        self.lbl_prompt.AutoSize = False
        self.Controls.Add(self.lbl_prompt)

        self.search_box = TextBox()
        self.search_box.Top = 55
        self.search_box.Left = 20
        self.search_box.Width = 780
        self.search_box.TextChanged += self._on_search_changed
        self.Controls.Add(self.search_box)

        self.tree_view = TreeView()
        self.tree_view.Top = 85
        self.tree_view.Left = 20
        self.tree_view.Width = 780
        self.tree_view.Height = 590
        self.tree_view.HideSelection = False
        self.tree_view.CheckBoxes = True
        self.tree_view.AfterCheck += self._on_after_check
        self.Controls.Add(self.tree_view)

        self.lbl_selected = Label()
        self.lbl_selected.Text = "Selected: none"
        self.lbl_selected.Top = 680
        self.lbl_selected.Left = 20
        self.lbl_selected.Width = 580
        self.Controls.Add(self.lbl_selected)

        btn_ok = Button()
        btn_ok.Text = "Select"
        btn_ok.Top = 675
        btn_ok.Left = 640
        btn_ok.Width = 75
        btn_ok.DialogResult = DialogResult.OK
        btn_ok.Click += self._on_ok_clicked
        self.Controls.Add(btn_ok)
        self.AcceptButton = btn_ok

        btn_cancel = Button()
        btn_cancel.Text = "Cancel"
        btn_cancel.Top = 675
        btn_cancel.Left = 725
        btn_cancel.Width = 75
        btn_cancel.DialogResult = DialogResult.Cancel
        self.Controls.Add(btn_cancel)
        self.CancelButton = btn_cancel

        self._build_tree(None)

    def _build_tree(self, search_filter):
        self.tree_view.Nodes.Clear()
        search = (search_filter or "").strip().lower()

        for group_key in sorted(self.catalog.keys(), key=natural_sort_key):
            labels = self.catalog[group_key]
            total_group_count = sum(len(rows) for rows in labels.values())
            group_node = TreeNode("{} ({})".format(
                group_key, total_group_count))
            group_node.Tag = ("group", group_key)

            for label in sorted(labels.keys(), key=natural_sort_key):
                rows = labels[label]
                label_node = TreeNode("{} ({})".format(label, len(rows)))
                label_node.Tag = ("label", group_key, label)

                for row in rows:
                    family_name = row["family"]
                    type_name = row["type"]
                    elem_id_int = row["elem_id_int"]
                    elevation = row["elevation"]

                    haystack = "{} {} {} {} {}".format(
                        group_key, label, family_name, type_name, elem_id_int
                    ).lower()
                    if search and search not in haystack:
                        continue

                    item_text = "{} | {} | ID {} | Elev: {}".format(
                        family_name, type_name, elem_id_int, elevation)
                    item_node = TreeNode(item_text)
                    item_node.Tag = ("item", elem_id_int, label,
                                     family_name, type_name, elevation)
                    label_node.Nodes.Add(item_node)

                if label_node.Nodes.Count > 0:
                    group_node.Nodes.Add(label_node)

            if group_node.Nodes.Count > 0:
                self.tree_view.Nodes.Add(group_node)

        self.tree_view.CollapseAll()

    def _on_search_changed(self, sender, args):
        self._build_tree(sender.Text)

    def _on_after_check(self, sender, args):
        node = args.Node
        if node is None:
            return

        self.tree_view.AfterCheck -= self._on_after_check
        try:
            if node.Tag and node.Tag[0] == "group":
                for label_node in node.Nodes:
                    label_node.Checked = node.Checked
                    for item_node in label_node.Nodes:
                        item_node.Checked = node.Checked
            elif node.Tag and node.Tag[0] == "label":
                for child in node.Nodes:
                    child.Checked = node.Checked
        finally:
            self.tree_view.AfterCheck += self._on_after_check

        checked_items = self.get_checked_item_nodes()
        if len(checked_items) == 0:
            self.lbl_selected.Text = "Selected: none"
        elif len(checked_items) == 1:
            _, _, label, family_name, type_name, elevation = checked_items[0].Tag
            self.lbl_selected.Text = "Selected: {} | {} | {} | Elev: {}".format(
                label, family_name, type_name, elevation)
        else:
            self.lbl_selected.Text = "Selected: {} equipment nodes".format(
                len(checked_items))

    def get_checked_item_nodes(self):
        checked = []
        for group_node in self.tree_view.Nodes:
            for label_node in group_node.Nodes:
                for item_node in label_node.Nodes:
                    if item_node.Checked and item_node.Tag and item_node.Tag[0] == "item":
                        checked.append(item_node)
        return checked

    def _on_ok_clicked(self, sender, args):
        checked_items = self.get_checked_item_nodes()
        if len(checked_items) == 0:
            TaskDialog.Show(
                "Select Equipment", "Check at least one equipment node before clicking Select.")
            self.DialogResult = getattr(DialogResult, "None")
            return

        self.selected_element_ids = []
        self.selected_items = []
        for node in checked_items:
            _, elem_id_int, label, family_name, type_name, elevation = node.Tag
            self.selected_element_ids.append(int(elem_id_int))
            self.selected_items.append({
                "elem_id_int": int(elem_id_int),
                "label": label,
                "family": family_name,
                "type": type_name,
                "elevation": elevation,
            })


try:
    equipment = collect_view_equipment_instances(doc, view)
    if not equipment:
        TaskDialog.Show(
            "No Equipment", "No Mechanical Equipment elements were found in the active view.")
        script.exit()

    catalog = build_bbm_catalog(equipment, doc)

    picker = EquipmentByLabelPicker(
        "Find Equipment by BBM Label",
        "Select one or more equipment items from the active view, grouped by BBM label.",
        catalog,
    )

    if picker.ShowDialog() != DialogResult.OK:
        script.exit()

    if not picker.selected_element_ids:
        script.exit()

    selected_ids = List[ElementId]()
    for elem_id_int in picker.selected_element_ids:
        selected_ids.Add(ElementId(elem_id_int))
    uidoc.Selection.SetElementIds(selected_ids)

    selected_by_id = {}
    for row in picker.selected_items:
        selected_by_id[row["elem_id_int"]] = row

    sorted_selected_ids = sorted(selected_by_id.keys())
    all_selected_eids = List[ElementId]()
    for elem_id_int in sorted_selected_ids:
        all_selected_eids.Add(ElementId(elem_id_int))

    output.print_md("# Equipment selected by BBM label")
    output.print_md(
        "- Selected count: {}".format(len(picker.selected_element_ids)))

    output.print_md("## Index")
    for i, elem_id_int in enumerate(sorted_selected_ids, start=1):
        row = selected_by_id[elem_id_int]
        output.print_md("{} | {} | {} | Elev: {}".format(
            format_index(i),
            row["label"],
            output.linkify(ElementId(elem_id_int)),
            row["elevation"],
        ))

    output.print_md("## Actions")
    output.print_md("- {}".format(output.linkify(all_selected_eids,
                    title="Select All Listed Equipment")))

    print_disclaimer(output)

except Exception as exc:
    TaskDialog.Show("Error", "Script failed: {}".format(str(exc)))
    import traceback

    output.print_md("```\n{}\n```".format(traceback.format_exc()))
