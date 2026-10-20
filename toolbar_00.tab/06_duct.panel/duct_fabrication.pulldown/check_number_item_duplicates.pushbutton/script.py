# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, ElementId
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent, ExternalEventRequest
from pyrevit import revit, script
from System.Collections.Generic import List
from System.Drawing import Size, Color, Font, FontStyle
from System.Windows.Forms import (
    AnchorStyles, Button, Form, FormStartPosition, Label, MessageBox,
    MessageBoxButtons, MessageBoxIcon, TextBox, TreeNode, TreeView,
)
from config.parameters_registry import *
import re
import traceback

import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")


# Button info
# ===================================================
__title__ = "Check Item Numbers"
__doc__ = """
Search duplicate item-number groups in the active view and select checked ducts.
Double-click a duct ID to select it and zoom to it without closing the menu."""
__persistentengine__ = True
__fullframeengine__ = True

# Variables
# ==================================================
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
view = revit.active_view
output = script.get_output()

# Helpers
# ========================================================================


def element_id_value(eid):
    if eid is None:
        return None
    return eid.IntegerValue if hasattr(eid, "IntegerValue") else eid.Value


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
            return element_id_value(param.AsElementId())
    except Exception:
        return None


def lookup_parameter_case_insensitive(element, param_name):
    """Case-insensitive parameter lookup"""
    param_name_lower = param_name.strip().lower()
    for param in element.Parameters:
        if param.Definition.Name.strip().lower() == param_name_lower:
            return param
    return None


def natural_sort_key(value):
    return [int(part) if part.isdigit() else part.lower()
            for part in re.split(r"(\d+)", value)]


class SelectDuctsHandler(IExternalEventHandler):
    def __init__(self, document):
        self.document = document
        self.element_ids = []
        self.pending = False
        self.zoom = False
        self.id_list_type = List[ElementId]
        self.message_box = MessageBox
        self.message_buttons = MessageBoxButtons
        self.message_icons = MessageBoxIcon
        self.format_traceback = traceback.format_exc

    def Execute(self, uiapp):
        try:
            active_uidoc = uiapp.ActiveUIDocument
            if (active_uidoc is None or not self.document.IsValidObject
                    or not active_uidoc.Document.Equals(self.document)):
                raise RuntimeError("Activate the original document before selecting ducts.")
            if not self.element_ids:
                return
            duct_ids = self.id_list_type()
            for element_id in self.element_ids:
                if self.document.GetElement(element_id) is None:
                    raise RuntimeError(
                        "A checked duct no longer exists. Close and reopen this menu.")
                duct_ids.Add(element_id)
            active_uidoc.Selection.SetElementIds(duct_ids)
            if self.zoom:
                active_uidoc.ShowElements(duct_ids)
        except Exception as ex:
            self.message_box.Show(
                "Selection failed:\n{}".format(ex), "Check Item Numbers",
                self.message_buttons.OK, self.message_icons.Error)
            print(self.format_traceback())
        finally:
            self.element_ids = []
            self.pending = False
            self.zoom = False

    def GetName(self):
        return "Select Duplicate Item Number Ducts"


class DuplicateItemForm(Form):
    def __init__(self, groups, select_event, select_handler):
        Form.__init__(self)
        self.Text = "Select Ducts by Duplicate Item Number"
        self.ClientSize = Size(700, 550)
        self.MinimumSize = Size(660, 400)
        self.StartPosition = FormStartPosition.CenterScreen
        self.groups = groups
        self.select_event = select_event
        self.select_handler = select_handler
        self.checked_ids = set()
        self.updating = False
        self.message_box = MessageBox
        self.message_buttons = MessageBoxButtons
        self.message_icons = MessageBoxIcon
        self.accepted_request = ExternalEventRequest.Accepted
        self.format_traceback = traceback.format_exc
        self.element_ids = {
            element_id_value(eid): eid for ids in groups.values() for eid in ids
        }
        self.link_font = Font(self.Font, FontStyle.Underline)

        search_label = Label()
        search_label.Text = "Search item number:"
        search_label.SetBounds(12, 15, 140, 24)
        self.Controls.Add(search_label)

        self.search_box = TextBox()
        self.search_box.SetBounds(152, 12, 536, 24)
        self.search_box.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.search_box.TextChanged += self._filter_tree
        self.Controls.Add(self.search_box)

        self.tree_view = TreeView()
        self.tree_view.SetBounds(12, 48, 676, 414)
        self.tree_view.CheckBoxes = True
        self.tree_view.Anchor = (
            AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right)
        self.tree_view.AfterCheck += self._on_node_checked
        self.tree_view.NodeMouseDoubleClick += self._on_id_double_click
        self.tree_view.ShowNodeToolTips = True
        self.Controls.Add(self.tree_view)

        self.summary = Label()
        self.summary.SetBounds(12, 474, 676, 24)
        self.summary.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.Controls.Add(self.summary)

        for text, left, width, callback in (
                ("Select All", 12, 120, self._on_select_all),
                ("Select Checked", 140, 150, self._on_select_checked),
                ("Deselect All", 298, 120, self._on_deselect_all),
                ("Close", 426, 100, self._on_close)):
            button = Button()
            button.Text = text
            button.SetBounds(left, 508, width, 30)
            button.Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            button.Click += callback
            self.Controls.Add(button)

        self.FormClosed += self._on_form_closed
        self._build_tree()

    def _build_tree(self):
        search = self.search_box.Text.strip().lower()
        self.updating = True
        self.tree_view.BeginUpdate()
        try:
            self.tree_view.Nodes.Clear()
            for item_number in sorted(self.groups, key=natural_sort_key):
                if search and search not in item_number.lower():
                    continue
                elements = self.groups[item_number]
                parent = TreeNode(
                    "Item Number: {} ({} ducts)".format(item_number, len(elements)))
                parent.Tag = ("group", item_number)
                for element_id in sorted(elements, key=element_id_value):
                    id_value = element_id_value(element_id)
                    child = TreeNode("ID: {}".format(id_value))
                    child.Tag = ("element", id_value)
                    child.ForeColor = Color.Blue
                    child.NodeFont = self.link_font
                    child.ToolTipText = "Double-click to select and zoom to this duct"
                    child.Checked = id_value in self.checked_ids
                    parent.Nodes.Add(child)
                parent.Checked = all(child.Checked for child in parent.Nodes)
                self.tree_view.Nodes.Add(parent)
        finally:
            self.tree_view.EndUpdate()
            self.updating = False
        self._update_summary()

    def _update_summary(self):
        self.summary.Text = "{} duplicate groups | {} ducts | {} checked".format(
            len(self.groups), sum(len(ids) for ids in self.groups.values()),
            len(self.checked_ids))

    def _filter_tree(self, sender, args):
        self._build_tree()

    def _on_node_checked(self, sender, args):
        if self.updating:
            return
        self.updating = True
        try:
            node = args.Node
            children = node.Nodes if node.Tag[0] == "group" else [node]
            for child in children:
                child.Checked = node.Checked
                id_value = child.Tag[1]
                if child.Checked:
                    self.checked_ids.add(id_value)
                else:
                    self.checked_ids.discard(id_value)
            if node.Parent is not None:
                node.Parent.Checked = all(child.Checked for child in node.Parent.Nodes)
        finally:
            self.updating = False
        self._update_summary()

    def _on_select_all(self, sender, args):
        # Select All applies to the groups visible under the current search.
        for parent in self.tree_view.Nodes:
            for child in parent.Nodes:
                self.checked_ids.add(child.Tag[1])
        self._build_tree()

    def _on_deselect_all(self, sender, args):
        self.checked_ids.clear()
        self._build_tree()

    def _on_select_checked(self, sender, args):
        if not self.checked_ids:
            self.message_box.Show(
                "Check at least one group or duct first.", self.Text,
                self.message_buttons.OK, self.message_icons.Information)
            return
        self._request_selection([
            self.element_ids[id_value] for id_value in sorted(self.checked_ids)
        ])

    def _on_id_double_click(self, sender, args):
        node = args.Node
        if node.Tag is not None and node.Tag[0] == "element":
            self._request_selection([self.element_ids[node.Tag[1]]], zoom=True)

    def _request_selection(self, element_ids, zoom=False):
        if self.select_handler.pending:
            self.message_box.Show(
                "A selection request is already pending. Please wait.", self.Text,
                self.message_buttons.OK, self.message_icons.Information)
            return
        try:
            self.select_handler.element_ids = element_ids
            self.select_handler.zoom = zoom
            self.select_handler.pending = True
            result = self.select_event.Raise()
            if result != self.accepted_request:
                raise RuntimeError("Revit did not accept the selection request: {}".format(result))
        except Exception as ex:
            self.select_handler.element_ids = []
            self.select_handler.pending = False
            self.select_handler.zoom = False
            self.message_box.Show(
                "Selection request failed:\n{}".format(ex), self.Text,
                self.message_buttons.OK, self.message_icons.Error)
            print(self.format_traceback())

    def _on_close(self, sender, args):
        self.Close()

    def _on_form_closed(self, sender, args):
        self.select_handler.element_ids = []
        self.select_event.Dispose()
        for parent in self.tree_view.Nodes:
            for child in parent.Nodes:
                child.NodeFont = None
        self.tree_view.Nodes.Clear()
        self.link_font.Dispose()


skip_values = {
    'skip',
    '0',
}

check_parameters = {
    RVT_ITEM_NUMBER,
}

# Main Code
# ==================================================
try:
    # Collect fabrication ductwork in the active view
    fab_duct = (FilteredElementCollector(doc, view.Id)
                .OfCategory(BuiltInCategory.OST_FabricationDuctwork)
                .WhereElementIsNotElementType()
                .ToElements())

    all_duct = list(fab_duct)
    if not all_duct:
        output.print_md("## No fabrication ducts found in view")
        script.exit()

    # Build parameter -> value -> elements map
    # Key is combination of all checked parameters
    param_groups = {}

    for d in all_duct:
        param_values = {}
        skip_element = False

        # Get values for all parameters in check_parameters
        for param_name in check_parameters:
            param = lookup_parameter_case_insensitive(d, param_name)
            value = get_param_value(param) if param else None

            # Clean and normalize the value
            if value is not None:
                value_str = str(value).strip().lower()
            else:
                value_str = None

            param_values[param_name] = value_str

            # Skip if empty or in skip_values
            empty_vals = {None, "", "**"}
            if value_str in empty_vals:
                skip_element = True
                break
            if value_str in skip_values:
                skip_element = True
                break

        if skip_element:
            continue

        # Group by combination of all parameters (use cleaned values)
        composite_key = " | ".join(str(param_values[p]).strip(
        ) if param_values[p] else "" for p in sorted(check_parameters))
        if composite_key not in param_groups:
            param_groups[composite_key] = []
        param_groups[composite_key].append(d)

    if not param_groups:
        output.print_md(
            "## No ducts found with parameters populated: {}".format(", ".join(check_parameters)))
        script.exit()

    duplicate_groups = {
        key: [duct.Id for duct in ducts]
        for key, ducts in param_groups.items() if len(ducts) > 1
    }
    if not duplicate_groups:
        output.print_md("## No matching ducts found")
        script.exit()

    select_handler = SelectDuctsHandler(doc)
    select_event = ExternalEvent.Create(select_handler)
    duplicate_form = DuplicateItemForm(duplicate_groups, select_event, select_handler)
    duplicate_form.Show()

except Exception as e:
    output.print_md("## Error: {}".format(str(e)))
    print(traceback.format_exc())
