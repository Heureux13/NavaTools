# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

from pyrevit import script, revit, forms, DB
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BoundingBoxXYZ,
    BuiltInCategory,
    ElementId,
    Category,
    View3D,
    ViewSection,
    ViewFamily,
    ViewFamilyType,
    XYZ,
    Transform,
    View,
)
from System.Collections.Generic import List
from System.Windows.Forms import (
    Form,
    Button,
    DialogResult,
    TextBox,
    TreeView,
    TreeNode,
    CheckBox,
    FormStartPosition,
)

# Variables
# =======================================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument


class WindowView


class EquipmentSelectionForm(Form):
    def __init__(self, family_names):
        Form.__init__(self)
        self.Text = "Select Mechanical Equipment"
        self.Width = 700
        self.Height = 630
        self.StartPosition = FormStartPosition.CenterScreen

        self.family_names = family_names
        self.hierarchy = self._build_hierarchy(family_names)
        self.checked_families = set()
        self._suppress_after_check = False

        self.search_box = TextBox()
        self.search_box.Top = 10
        self.search_box.Left = 10
        self.search_box.Width = 660
        self.search_box.TextChanged += self._filter_tree
        self.Controls.Add(self.search_box)

        self.tree_view = TreeView()
        self.tree_view.Top = 40
        self.tree_view.Left = 10
        self.tree_view.Width = 660
        self.tree_view.Height = 380
        self.tree_view.CheckBoxes = True
        self.tree_view.AfterCheck += self._on_node_checked
        self.Controls.Add(self.tree_view)

        self.create_north_checkbox = CheckBox()
        self.create_north_checkbox.Text = "Create horizontal section (North)"
        self.create_north_checkbox.Top = 430
        self.create_north_checkbox.Left = 10
        self.create_north_checkbox.Width = 320
        self.create_north_checkbox.Checked = False
        self.Controls.Add(self.create_north_checkbox)

        self.create_east_checkbox = CheckBox()
        self.create_east_checkbox.Text = "Create vertical section (East)"
        self.create_east_checkbox.Top = 455
        self.create_east_checkbox.Left = 10
        self.create_east_checkbox.Width = 320
        self.create_east_checkbox.Checked = False
        self.Controls.Add(self.create_east_checkbox)

        self.create_3d_checkbox = CheckBox()
        self.create_3d_checkbox.Text = "Create 3D view"
        self.create_3d_checkbox.Top = 480
        self.create_3d_checkbox.Left = 10
        self.create_3d_checkbox.Width = 320
        self.create_3d_checkbox.Checked = False
        self.Controls.Add(self.create_3d_checkbox)

        btn_all = Button()
        btn_all.Text = "Select All"
        btn_all.Top = 520
        btn_all.Left = 10
        btn_all.Width = 120
        btn_all.Click += self._on_select_all
        self.Controls.Add(btn_all)

        btn_select = Button()
        btn_select.Text = "Create Views"
        btn_select.Top = 520
        btn_select.Left = 140
        btn_select.Width = 140
        btn_select.DialogResult = DialogResult.Yes
        self.Controls.Add(btn_select)
        self.AcceptButton = btn_select

        btn_cancel = Button()
        btn_cancel.Text = "Cancel"
        btn_cancel.Top = 520
        btn_cancel.Left = 290
        btn_cancel.Width = 120
        btn_cancel.DialogResult = DialogResult.Cancel
        self.Controls.Add(btn_cancel)
        self.CancelButton = btn_cancel

        self._build_tree()
        self.tree_view.CollapseAll()

    def _build_hierarchy(self, names):
        hierarchy = {}
        for full_name in names:
            base_name = full_name.split("(")[0].strip()
            if base_name not in hierarchy:
                hierarchy[base_name] = []
            hierarchy[base_name].append(full_name)
        return hierarchy

    def _build_tree(self, search_filter=None):
        self.tree_view.Nodes.Clear()

        for base_name in sorted(self.hierarchy.keys()):
            variants = self.hierarchy[base_name]
            base_matches = (not search_filter) or (
                search_filter in base_name.lower())
            variant_matches = [v for v in variants if (
                not search_filter) or (search_filter in v.lower())]

            if not base_matches and not variant_matches:
                continue

            parent_node = TreeNode(base_name)
            parent_node.Tag = ("parent", base_name)

            for variant in sorted(variants):
                if (not search_filter) or (search_filter in variant.lower()):
                    child_node = TreeNode(variant)
                    child_node.Tag = ("child", variant)
                    if variant in self.checked_families:
                        child_node.Checked = True
                    parent_node.Nodes.Add(child_node)

            for child_node in parent_node.Nodes:
                if child_node.Checked:
                    parent_node.Checked = True
                    break

            if parent_node.Nodes.Count > 0:
                self.tree_view.Nodes.Add(parent_node)

    def _filter_tree(self, sender, args):
        search = sender.Text.lower().strip()
        self._build_tree(search if search else None)
        self.tree_view.CollapseAll()

    def _on_select_all(self, sender, ):
        self._suppress_after_check = True
        try:
            for parent_node in self.tree_view.Nodes:
                parent_node.Checked = True
                for child_node in parent_node.Nodes:
                    child_node.Checked = True
                    if child_node.Tag and child_node.Tag[0] == "child":
                        self.checked_families.add(child_node.Tag[1])
        finally:
            self._suppress_after_check = False

    def _on_node_checked(self, sender, args):
        if self._suppress_after_check:
            return

        node = args.Node
        if not node or not node.Tag:
            return

        self._suppress_after_check = True
        try:
            kind, value = node.Tag

            if kind == "parent":
                for child_node in node.Nodes:
                    child_node.Checked = node.Checked
                    if child_node.Tag and child_node.Tag[0] == "child":
                        child_variant = child_node.Tag[1]
                        if node.Checked:
                            self.checked_families.add(child_variant)
                        else:
                            self.checked_families.discard(child_variant)

            elif kind == "child":
                if node.Checked:
                    self.checked_families.add(value)
                else:
                    self.checked_families.discard(value)
                    if node.Parent:
                        parent_checked = any(
                            c.Checked for c in node.Parent.Nodes)
                        node.Parent.Checked = parent_checked
        finally:
            self._suppress_after_check = False

    def get_selected_families(self):
        # Read from visible tree and merge with tracked checked state.
        selected = set(self.checked_families)
        for parent_node in self.tree_view.Nodes:
            for child_node in parent_node.Nodes:
                if child_node.Checked and child_node.Tag and child_node.Tag[0] == "child":
                    selected.add(child_node.Tag[1])
        return sorted(selected)

    def get_create_north_section(self):
        return self.create_north_checkbox.Checked

    def get_create_east_section(self):
        return self.create_east_checkbox.Checked

    def get_create_3d_view(self):
        return self.create_3d_checkbox.Checked


def get_family_name(elem):
    symbol = getattr(elem, "Symbol", None)
    if symbol and getattr(symbol, "Family", None):
        return symbol.Family.Name

    fam = getattr(elem, "Family", None)
    if fam:
        return fam.Name

    return None
