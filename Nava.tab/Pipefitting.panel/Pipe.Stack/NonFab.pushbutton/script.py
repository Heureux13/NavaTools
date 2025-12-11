# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, ElementId, BoundingBoxIntersectsFilter, Outline
from Autodesk.Revit.UI import TaskDialog
from pyrevit import revit, script
from System.Windows.Forms import Form, Label, Button, DialogResult, TextBox, TreeView, TreeNode
from System.Collections.Generic import List
from revit_duct import RevitDuct
from revit_output import print_disclaimer
import clr
import re
clr.AddReference("System.Windows.Forms")


# Button info
# ===================================================
__title__ = "Non-Fab"
__doc__ = """
Selects all non-fabrication piping, and filters them down by parameters to select
"""

# Variables
# ==================================================
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
view = revit.active_view
output = script.get_output()

# Class
# =====================================================================


class EnhancedParamForm(Form):
    def __init__(self, param_groups):
        Form.__init__(self)
        self.Text = "Select Pipes by Parameter"
        self.Width = 700
        self.Height = 700
        self.param_groups = param_groups

        # Search box
        self.search_box = TextBox()
        self.search_box.Top = 20
        self.search_box.Left = 20
        self.search_box.Width = 640
        self.search_box.PlaceholderText = "Search parameters..."
        self.search_box.TextChanged += self._filter_tree
        self.Controls.Add(self.search_box)

        # TreeView with checkboxes for expandable hierarchy
        self.tree_view = TreeView()
        self.tree_view.Top = 50
        self.tree_view.Left = 20
        self.tree_view.Width = 640
        self.tree_view.Height = 550
        self.tree_view.CheckBoxes = True
        self.tree_view.AfterCheck += self._on_node_checked
        self.Controls.Add(self.tree_view)

        # Build tree structure
        self._build_tree()

        # OK button
        btn_ok = Button()
        btn_ok.Text = "Select All Checked"
        btn_ok.Top = 620
        btn_ok.Left = 20
        btn_ok.Width = 200
        btn_ok.DialogResult = DialogResult.OK
        self.Controls.Add(btn_ok)
        self.AcceptButton = btn_ok

    def _build_tree(self, search_filter=None):
        self.tree_view.Nodes.Clear()

        for param_name in sorted(self.param_groups.keys(), key=natural_sort_key):
            # Check if parameter name matches search
            param_matches = not search_filter or search_filter in param_name.lower()

            # Create parent node for parameter
            param_node = TreeNode(param_name)
            param_node.Tag = ("param", param_name)

            # Add child nodes for each value
            for value in sorted(self.param_groups[param_name].keys(), key=natural_sort_key):
                # If parameter name matches, show all its values
                # Otherwise, only show values that match the search
                if param_matches or (search_filter and search_filter in str(value).lower()):
                    value_text = "{} ({} parts)".format(
                        value, len(self.param_groups[param_name][value]))
                    value_node = TreeNode(value_text)
                    value_node.Tag = ("value", param_name, value)
                    param_node.Nodes.Add(value_node)

            # Only add param node if it has children
            if param_node.Nodes.Count > 0:
                self.tree_view.Nodes.Add(param_node)

    def _filter_tree(self, sender, args):
        search = sender.Text.lower()
        self._build_tree(search if search else None)

    def _on_node_checked(self, sender, args):
        # Prevent event recursion
        self.tree_view.AfterCheck -= self._on_node_checked

        node = args.Node
        # If parent is checked/unchecked, check/uncheck all children
        if node.Tag and node.Tag[0] == "param":
            for child in node.Nodes:
                child.Checked = node.Checked

        self.tree_view.AfterCheck += self._on_node_checked

    def get_checked_ducts(self):
        """Returns list of duct elements from all checked nodes"""
        ducts = set()

        for param_node in self.tree_view.Nodes:
            # If parent is checked, add all ducts from this parameter
            if param_node.Checked and param_node.Tag[0] == "param":
                param_name = param_node.Tag[1]
                for value_list in self.param_groups[param_name].values():
                    for duct in value_list:
                        ducts.add(duct)
            else:
                # Check individual value nodes
                for value_node in param_node.Nodes:
                    if value_node.Checked and value_node.Tag[0] == "value":
                        param_name = value_node.Tag[1]
                        value = value_node.Tag[2]
                        for duct in self.param_groups[param_name][value]:
                            ducts.add(duct)

        return list(ducts)

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
        # Prefer string if present
        if param.AsString():
            return param.AsString()
        # Next, use value string (formatted)
        if param.AsValueString():
            return param.AsValueString()
        # Finally, raw numeric
        if param.StorageType == 1:  # Double
            return param.AsDouble()
        if param.StorageType == 2:  # Integer
            return param.AsInteger()
        if param.StorageType == 3:  # ElementId
            return param.AsElementId().IntegerValue
    except Exception:
        return None
    return None


# Main Code
# ==================================================
try:
    # Get view's crop box to filter elements within view range
    crop_box = view.CropBox
    outline = Outline(crop_box.Min, crop_box.Max)
    bbox_filter = BoundingBoxIntersectsFilter(outline)

    all_straights = (FilteredElementCollector(doc, view.Id)
                     .OfCategory(BuiltInCategory.OST_PipeCurves)
                     .WhereElementIsNotElementType()
                     .WherePasses(bbox_filter)
                     .ToElements()
                     )

    all_fittings = (FilteredElementCollector(doc, view.Id)
                    .OfCategory(BuiltInCategory.OST_PipeFitting)
                    .WhereElementIsNotElementType()
                    .WherePasses(bbox_filter)
                    .ToElements()
                    )

    # Filter to only elements visible in view (not hidden)
    all_straights = [e for e in all_straights if not e.IsHidden(view)]
    all_fittings = [e for e in all_fittings if not e.IsHidden(view)]

    # Combines both list into one
    all_duct = list(all_straights) + list(all_fittings)

    # Debug: verify view filtering worked
    output.print_md("**Debug:** Found {} pipes in view '{}'".format(len(all_duct), view.Name))

    # Build parameter -> value -> elements map
    param_groups = {}
    for d in all_duct:
        try:
            for p in list(d.Parameters):
                pname = p.Definition.Name
                pval = get_param_value(p)
                if pval is None or pval == "":
                    continue
                if pname not in param_groups:
                    param_groups[pname] = {}
                if pval not in param_groups[pname]:
                    param_groups[pname][pval] = []
                param_groups[pname][pval].append(d)
        except Exception as e:
            output.print_md("Error reading parameters: {}".format(str(e)))
            continue

    if not param_groups:
        TaskDialog.Show(
            "No Parameters",
            "No parameter data found on pipes in view."
        )
        script.exit()

    # Show window with checkbox parameter selection
    form = EnhancedParamForm(param_groups)
    result = form.ShowDialog()

    # Stop script if user does not select OK
    if result != DialogResult.OK:
        script.exit()

    duct_run = form.get_checked_ducts()
    if not duct_run:
        TaskDialog.Show("No Selection", "No pipes were selected.")
        script.exit()

    # Select ducts in Revit
    duct_ids = List[ElementId]()
    for d in duct_run:
        duct_ids.Add(d.Id)
    uidoc.Selection.SetElementIds(duct_ids)

    # final printout with links to duct
    if len(duct_run) < 500:
        for i, d in enumerate(duct_run, start=1):
            duct_obj = RevitDuct(doc, view, d)
            family_name = duct_obj.family if duct_obj.family else "Unknown"
            output.print_md(
                "### No: {} | ID: {} | Family: {} | View: {}".format(
                    i,
                    output.linkify(d.Id),
                    family_name,
                    view.Name
                )
            )

    element_ids = [d.Id for d in duct_run]
    output.print_md("---")
    output.print_md(
        "# Total Elements: {:03}, {}".format(
            len(duct_run),
            output.linkify(element_ids)
        )
    )

    # Final print statements
    print_disclaimer(output)

except Exception as e:
    TaskDialog.Show("Error", "Script failed: {}".format(str(e)))
    import traceback
    output.print_md("```\n{}\n```".format(traceback.format_exc()))
