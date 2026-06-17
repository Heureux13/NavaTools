# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    ElementId,
    VisibleInViewFilter,
    Transaction,
)
from pyrevit import revit, script, forms
from System.Windows.Forms import (
    Form,
    Button,
    DialogResult,
    TextBox,
    TreeView,
    TreeNode,
    CheckBox,
    FormStartPosition,
    AnchorStyles,
)
from System.Collections.Generic import List
from System.Drawing import Size
import re
import os


# Button info
# ===================================================
__title__ = "Select by Fab Note"
__doc__ = """
Select by Fab Notes with Item Number
"""

# Variables
# ==================================================
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
view = revit.active_view
output = script.get_output()

families_to_skip = {
    "spiral duct",
    "duct spiral"
    "boot tap wdamper",
    "boot saddle tap",
    "boot tap - wdamper",
    "coupling",
}

# Values to skip in Fabrication Notes
fab_notes_to_skip = {
    "0",
    "skip",
}

# Class
# =====================================================================


class EnhancedParamForm(Form):
    def __init__(self, param_groups, uidoc):
        Form.__init__(self)
        self.Text = "Select Ducts by Fabrication Notes"
        self.Width = 700
        self.Height = 600
        self.MinimumSize = Size(500, 400)
        self.StartPosition = FormStartPosition.CenterScreen
        self.param_groups = param_groups
        self.uidoc = uidoc

        # Build hierarchical structure: base_name -> [variants]
        self.hierarchy = self._build_hierarchy(param_groups)

        _MARGIN = 12
        _BTN_H = 30
        _BTN_TOP = self.ClientSize.Height - _MARGIN - _BTN_H
        _SEARCH_H = 24
        _inner_w = self.ClientSize.Width - _MARGIN * 2

        # Search box
        self.search_box = TextBox()
        self.search_box.Top = _MARGIN
        self.search_box.Left = _MARGIN
        self.search_box.Width = _inner_w
        self.search_box.Height = _SEARCH_H
        self.search_box.PlaceholderText = "Search..."
        self.search_box.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.search_box.TextChanged += self._filter_tree
        self.Controls.Add(self.search_box)

        # TreeView with checkboxes for expandable hierarchy
        _tree_top = _MARGIN + _SEARCH_H + _MARGIN
        self.tree_view = TreeView()
        self.tree_view.Top = _tree_top
        self.tree_view.Left = _MARGIN
        self.tree_view.Width = _inner_w
        self.tree_view.Height = _BTN_TOP - _tree_top - _MARGIN
        self.tree_view.CheckBoxes = True
        self.tree_view.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.tree_view.AfterCheck += self._on_node_checked
        self.Controls.Add(self.tree_view)

        # Build tree structure
        self._build_tree()

        # Select All button
        btn_all = Button()
        btn_all.Text = "Select All"
        btn_all.Top = _BTN_TOP
        btn_all.Left = _MARGIN
        btn_all.Width = 120
        btn_all.Height = _BTN_H
        btn_all.Anchor = AnchorStyles.Bottom | AnchorStyles.Left
        btn_all.Click += self._on_select_all
        self.Controls.Add(btn_all)

        # Select Checked button
        btn_itemize = Button()
        btn_itemize.Text = "Select Checked"
        btn_itemize.Top = _BTN_TOP
        btn_itemize.Left = _MARGIN + 120 + 8
        btn_itemize.Width = 150
        btn_itemize.Height = _BTN_H
        btn_itemize.Anchor = AnchorStyles.Bottom | AnchorStyles.Left
        btn_itemize.Click += self._on_select_checked
        self.Controls.Add(btn_itemize)

        # Close button
        btn_close = Button()
        btn_close.Text = "Close"
        btn_close.Top = _BTN_TOP
        btn_close.Left = _MARGIN + 120 + 8 + 150 + 8
        btn_close.Width = 100
        btn_close.Height = _BTN_H
        btn_close.Anchor = AnchorStyles.Bottom | AnchorStyles.Left
        btn_close.DialogResult = DialogResult.Cancel
        self.Controls.Add(btn_close)

    def _build_hierarchy(self, param_groups):
        """Build hierarchy by grouping variants under base names"""
        hierarchy = {}
        for full_name in param_groups.keys():
            # Extract base name (everything before parenthesis)
            base_name = full_name.split("(")[0].strip()
            if base_name not in hierarchy:
                hierarchy[base_name] = []
            hierarchy[base_name].append(full_name)
        return hierarchy

    def _build_tree(self, search_filter=None):
        """Build tree structure with parent-child relationships"""
        self.tree_view.Nodes.Clear()

        for base_name in sorted(self.hierarchy.keys(), key=natural_sort_key):
            variants = self.hierarchy[base_name]

            # Check if base name or any variant matches search
            base_matches = not search_filter or search_filter in base_name.lower()
            variant_matches = [
                v for v in variants if search_filter is None or search_filter in v.lower()]

            if not base_matches and not variant_matches:
                continue

            # Create parent node
            total_count = sum(len(self.param_groups.get(v, []))
                              for v in variants)
            parent_text = "{} ({} parts)".format(base_name, total_count)
            parent_node = TreeNode(parent_text)
            parent_node.Tag = ("parent", base_name)

            # Add child nodes for each variant
            for variant in sorted(variants, key=natural_sort_key):
                if search_filter is None or search_filter in variant.lower():
                    count = len(self.param_groups.get(variant, []))
                    child_text = "{} ({} parts)".format(variant, count)
                    child_node = TreeNode(child_text)
                    child_node.Tag = ("child", variant)
                    parent_node.Nodes.Add(child_node)

            if parent_node.Nodes.Count > 0:
                self.tree_view.Nodes.Add(parent_node)

    def _filter_tree(self, sender, args):
        """Filter tree based on search text"""
        search = sender.Text.lower()
        self._build_tree(search if search else None)

    def _on_select_all(self, sender, args):
        """Check all parent nodes to select all ducts"""
        self.tree_view.AfterCheck -= self._on_node_checked

        for parent_node in self.tree_view.Nodes:
            if parent_node.Tag and parent_node.Tag[0] == "parent":
                parent_node.Checked = True
                for child_node in parent_node.Nodes:
                    child_node.Checked = True

        self.tree_view.AfterCheck += self._on_node_checked

    def _on_select_checked(self, sender, args):
        """Select checked ducts in Revit without closing the form."""
        ducts = self.get_checked_ducts()
        if not ducts:
            return
        duct_ids = List[ElementId]()
        for d in ducts:
            duct_ids.Add(d.Id)
        self.uidoc.Selection.SetElementIds(duct_ids)

    def _on_node_checked(self, sender, args):
        """When a parent is checked/unchecked, check/uncheck all children"""
        self.tree_view.AfterCheck -= self._on_node_checked

        node = args.Node

        # If parent node is checked/unchecked, apply to all children
        if node.Tag and node.Tag[0] == "parent":
            for child_node in node.Nodes:
                child_node.Checked = node.Checked

        self.tree_view.AfterCheck += self._on_node_checked

    def get_checked_ducts(self):
        """Returns list of duct elements from all checked nodes"""
        ducts = set()

        for parent_node in self.tree_view.Nodes:
            if parent_node.Tag and parent_node.Tag[0] == "parent":
                # Check child nodes individually
                for child_node in parent_node.Nodes:
                    if child_node.Checked and child_node.Tag and child_node.Tag[0] == "child":
                        variant = child_node.Tag[1]
                        if variant in self.param_groups:
                            for duct in self.param_groups[variant]:
                                ducts.add(duct)

        return list(ducts)

    def get_checked_values(self):
        """Returns selected Fabrication Notes values from checked child nodes"""
        checked_values = []

        for parent_node in self.tree_view.Nodes:
            if parent_node.Tag and parent_node.Tag[0] == "parent":
                for child_node in parent_node.Nodes:
                    if child_node.Checked and child_node.Tag and child_node.Tag[0] == "child":
                        checked_values.append(child_node.Tag[1])

        return checked_values

# Helpers
# ========================================================================


def natural_sort_key(s):
    # Sort runs with natural/numeric sorting
    return [
        int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)
    ]


def get_param_value(param):
    try:
        if param.StorageType == 0:  # None
            return None
        if param.AsString():
            return param.AsString()
        if param.AsValueString():
            return param.AsValueString()
        if param.StorageType == 1:  # Double
            return param.AsDouble()
        if param.StorageType == 2:  # Integer
            return param.AsInteger()
        if param.StorageType == 3:  # ElementId
            return param.AsElementId().IntegerValue
    except Exception:
        return None


def is_param_true(param):
    if not param:
        return False
    try:
        # Integer/Yes-No parameters
        if param.StorageType == 2:
            return param.AsInteger() == 1
        val_str = param.AsString() or param.AsValueString() or ""
        return val_str.strip().lower() in ("true", "yes", "1")
    except Exception:
        return False


# Main Code
# ==================================================
try:
    # Collect only fabrication ductwork strictly visible in the active view
    fab_duct = (FilteredElementCollector(doc, view.Id)
                .OfCategory(BuiltInCategory.OST_FabricationDuctwork)
                .WhereElementIsNotElementType()
                .WherePasses(VisibleInViewFilter(doc, view.Id))
                .ToElements())

    all_duct = list(fab_duct)
    if not all_duct:
        script.exit()

    # Build parameter -> value -> elements map (only for ducts with numeric Item Number)
    param_groups = {}
    for d in all_duct:
        # Skip families in the skip list
        try:
            fam_param = d.LookupParameter("Family")
            if fam_param:
                fam_name = get_param_value(fam_param)
                if fam_name:
                    fam_lower = str(fam_name).strip().lower()
                    if any(skip_fam in fam_lower for skip_fam in families_to_skip):
                        continue
        except Exception:
            pass

        # First check if Item Number has a numeric value
        item_number_found = False
        for p in list(d.Parameters):
            if p.Definition.Name != "Item Number":
                continue
            item_val = get_param_value(p)
            if item_val is None or item_val == "":
                break  # No item number, skip this duct
            # Check if value is numeric
            item_str = str(item_val).replace(",", "").strip()
            try:
                num_val = float(item_str)
                # Skip if item number is 0
                if num_val == 0:
                    break
                item_number_found = True  # Has numeric Item Number
            except ValueError:
                break  # Item Number exists but not numeric, skip this duct
            break

        if not item_number_found:
            continue  # Skip ducts without numeric Item Number

        # Skip if Fab Exported is Yes
        fab_exported_param = d.LookupParameter("Fab Exported")
        if fab_exported_param and is_param_true(fab_exported_param):
            continue

        # Now get the Fabrication Notes value
        for p in list(d.Parameters):
            if p.Definition.Name != "Fabrication Notes":
                continue
            pval = get_param_value(p)
            if pval is None or pval == "":
                pval = "(blank)"
            else:
                # Keep full value with variants (don't strip parenthesis)
                pval = str(pval).strip()

            # Skip values in the skip list
            if pval.lower() in fab_notes_to_skip:
                break

            if pval not in param_groups:
                param_groups[pval] = []
            param_groups[pval].append(d)
            break

    if not param_groups:
        script.exit()

    # Show form — selection happens inside the form on each click
    form = EnhancedParamForm(param_groups, uidoc)
    form.ShowDialog()

except Exception as e:
    output.print_md("**Error:** {}".format(str(e)))
    import traceback
    output.print_md("```\n{}\n```".format(traceback.format_exc()))
