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
    CopyPasteOptions,
    ElementId,
    ElementTransformUtils,
    FilteredElementCollector,
    Transform,
    View,
    ViewSheet,
)
from System.Collections.Generic import List

# Button info
# ======================================================================
__title__ = 'Copy Room Tags'
__doc__ = '''
Select all matching Room Tags in active view,
then copy/paste them into selected views.
'''

# Variables
# ======================================================================

doc = revit.doc
uidoc = revit.uidoc
source_view = revit.active_view
output = script.get_output()


def _get_element_id_value(element_id):
    """Return ElementId value compatible with both Revit 2025 and 2026+."""
    if element_id is None:
        return None

    try:
        return element_id.Value
    except Exception:
        pass

    try:
        return element_id.IntegerValue
    except Exception:
        return None


class ListOption(object):
    def __init__(self, item, display_name):
        self.item = item
        self.display_name = display_name


def get_view_display_name(view):
    try:
        return '{} | {}'.format(view.ViewType, view.Name)
    except Exception:
        return str(view.Id)


def get_tag_family_and_type(tag):
    family_name = None
    type_name = None

    try:
        tag_type = doc.GetElement(tag.GetTypeId())
    except Exception:
        tag_type = None

    if tag_type:
        try:
            family_name = getattr(tag_type, 'FamilyName', None)
        except Exception:
            family_name = None

        try:
            type_name = getattr(tag_type, 'Name', None)
        except Exception:
            type_name = None

        if not family_name:
            try:
                fam_param = tag_type.get_Parameter(
                    BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
                if fam_param:
                    family_name = fam_param.AsString()
            except Exception:
                pass

        if not type_name:
            try:
                name_param = tag_type.get_Parameter(
                    BuiltInParameter.SYMBOL_NAME_PARAM)
                if name_param:
                    type_name = name_param.AsString()
            except Exception:
                pass

    return family_name, type_name


def collect_target_views(document, active):
    all_views = [
        v for v in FilteredElementCollector(document).OfClass(View)
        if not v.IsTemplate and not isinstance(v, ViewSheet)
    ]

    candidates = [
        v for v in all_views
        if v.Id != active.Id and v.ViewType == active.ViewType
    ]

    return sorted(candidates, key=lambda v: get_view_display_name(v).lower())


def collect_annotation_options(tags):
    """Group tags by family, then by type within each family."""
    by_family = {}
    for tag in tags:
        family_name, type_name = get_tag_family_and_type(tag)
        family_name = family_name or '(Unknown Family)'
        type_name = type_name or '(Unknown Type)'

        if family_name not in by_family:
            by_family[family_name] = {}
        if type_name not in by_family[family_name]:
            by_family[family_name][type_name] = []
        by_family[family_name][type_name].append(tag)

    return by_family


def select_tag_families_and_types(by_family):
    """Hierarchical TreeView selection with search and select all."""
    from System.Windows.Forms import (
        Form,
        Button,
        DialogResult,
        TreeView,
        TreeNode,
        TextBox,
        FormStartPosition,
        Label,
    )

    class TagSelectionForm(Form):
        def __init__(self, by_family):
            Form.__init__(self)
            self.Text = "Select Tag Families and Types"
            self.Width = 500
            self.Height = 650
            self.StartPosition = FormStartPosition.CenterScreen
            self.by_family = by_family
            self.all_nodes = []

            # Search label
            lbl_search = Label()
            lbl_search.Text = "Search:"
            lbl_search.Top = 10
            lbl_search.Left = 10
            lbl_search.Width = 50
            lbl_search.Height = 20
            self.Controls.Add(lbl_search)

            # Search box
            self.search_box = TextBox()
            self.search_box.Top = 10
            self.search_box.Left = 60
            self.search_box.Width = 410
            self.search_box.TextChanged += self._filter_tree
            self.Controls.Add(self.search_box)

            # TreeView with checkboxes
            self.tree_view = TreeView()
            self.tree_view.Top = 40
            self.tree_view.Left = 10
            self.tree_view.Width = 460
            self.tree_view.Height = 420
            self.tree_view.CheckBoxes = True
            self.tree_view.AfterCheck += self._on_node_checked
            self.Controls.Add(self.tree_view)

            # Build tree
            self._build_tree()

            # Select All button
            btn_all = Button()
            btn_all.Text = "Select All"
            btn_all.Top = 470
            btn_all.Left = 10
            btn_all.Width = 100
            btn_all.Click += self._select_all
            self.Controls.Add(btn_all)

            # Select Checked button
            btn_ok = Button()
            btn_ok.Text = "Select Checked"
            btn_ok.Top = 470
            btn_ok.Left = 320
            btn_ok.Width = 150
            btn_ok.DialogResult = DialogResult.OK
            self.Controls.Add(btn_ok)
            self.AcceptButton = btn_ok

            # Cancel button
            btn_cancel = Button()
            btn_cancel.Text = "Cancel"
            btn_cancel.Top = 470
            btn_cancel.Left = 480
            btn_cancel.Width = 80
            btn_cancel.DialogResult = DialogResult.Cancel
            self.Controls.Add(btn_cancel)
            self.CancelButton = btn_cancel

        def _build_tree(self):
            """Build tree with family parents and type children."""
            self.tree_view.Nodes.Clear()
            self.all_nodes = []

            for family_name in sorted(self.by_family.keys()):
                types_dict = self.by_family[family_name]
                total_count = sum(len(types_dict[t]) for t in types_dict)

                parent_text = "{} ({})".format(family_name, total_count)
                parent_node = TreeNode(parent_text)
                parent_node.Tag = ("family", family_name)

                for type_name in sorted(types_dict.keys()):
                    count = len(types_dict[type_name])
                    child_text = "{} ({})".format(type_name, count)
                    child_node = TreeNode(child_text)
                    child_node.Tag = (family_name, type_name)
                    parent_node.Nodes.Add(child_node)

                self.tree_view.Nodes.Add(parent_node)
                self.all_nodes.append(parent_node)

        def _filter_tree(self, sender, args):
            """Filter tree based on search text."""
            search = self.search_box.Text.lower()
            self.tree_view.Nodes.Clear()

            for parent_node in self.all_nodes:
                parent_text = parent_node.Text.lower()
                parent_matches = not search or search in parent_text

                # Check if any children match
                child_matches = []
                for child_node in parent_node.Nodes:
                    child_text = child_node.Text.lower()
                    if not search or search in child_text:
                        child_matches.append(child_node)

                if parent_matches or child_matches:
                    new_parent = TreeNode(parent_node.Text)
                    new_parent.Tag = parent_node.Tag
                    new_parent.Checked = parent_node.Checked

                    for child_node in child_matches:
                        new_child = TreeNode(child_node.Text)
                        new_child.Tag = child_node.Tag
                        new_child.Checked = child_node.Checked
                        new_parent.Nodes.Add(new_child)

                    self.tree_view.Nodes.Add(new_parent)

        def _select_all(self, sender, args):
            """Check all nodes in tree."""
            self.tree_view.AfterCheck -= self._on_node_checked
            for node in self.tree_view.Nodes:
                self._check_node_recursive(node, True)
            self.tree_view.AfterCheck += self._on_node_checked

        def _check_node_recursive(self, node, checked):
            """Recursively check/uncheck node and children."""
            node.Checked = checked
            for child in node.Nodes:
                self._check_node_recursive(child, checked)

        def _on_node_checked(self, sender, args):
            """When parent checked, check all children."""
            self.tree_view.AfterCheck -= self._on_node_checked

            node = args.Node
            if node.Tag and node.Tag[0] == "family":
                for child_node in node.Nodes:
                    child_node.Checked = node.Checked

            self.tree_view.AfterCheck += self._on_node_checked

        def get_checked_types(self):
            """Get all checked type keys."""
            result = set()

            def traverse(node):
                if node.Checked and node.Tag and node.Tag[0] != "family":
                    result.add(node.Tag)
                for child in node.Nodes:
                    traverse(child)

            for node in self.tree_view.Nodes:
                traverse(node)

            return result

    form = TagSelectionForm(by_family)
    if form.ShowDialog() == DialogResult.OK:
        return form.get_checked_types()
    return None


try:
    room_tags = list(
        FilteredElementCollector(doc, source_view.Id)
        .OfCategory(BuiltInCategory.OST_RoomTags)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    room_tags = [t for t in room_tags if t.OwnerViewId == source_view.Id]

    if not room_tags:
        output.print_md('## No room tags found in the active view.')
        script.exit()

    by_family = collect_annotation_options(room_tags)
    selected_keys = select_tag_families_and_types(by_family)
    if not selected_keys:
        script.exit()

    matching_tags = []
    for tag in room_tags:
        family_name, type_name = get_tag_family_and_type(tag)
        key = (family_name or '(Unknown Family)',
               type_name or '(Unknown Type)')
        if key in selected_keys:
            matching_tags.append(tag)

    if not matching_tags:
        output.print_md(
            '## No matching room tags found for the selected annotation(s).')
        script.exit()

    selected_ids = List[ElementId]([tag.Id for tag in matching_tags])
    uidoc.Selection.SetElementIds(selected_ids)

    target_views = collect_target_views(doc, source_view)
    if not target_views:
        output.print_md('## No compatible target views found.')
        script.exit()

    options = [ListOption(v, get_view_display_name(v)) for v in target_views]
    picked = forms.SelectFromList.show(
        options,
        name_attr='display_name',
        multiselect=True,
        title='Select Views to Paste Room Tags'
    )
    if not picked:
        script.exit()

    destination_views = [opt.item for opt in picked]
    copied_count = 0
    failed = []

    with revit.Transaction('Copy Room Tags To Views'):
        for destination_view in destination_views:
            try:
                ElementTransformUtils.CopyElements(
                    source_view,
                    selected_ids,
                    destination_view,
                    Transform.Identity,
                    CopyPasteOptions()
                )
                copied_count += 1
            except Exception as copy_error:
                failed.append((get_view_display_name(
                    destination_view), str(copy_error)))

    output.print_md(
        '### Selected {} room tags and pasted to {} view(s).'.format(
            len(matching_tags),
            copied_count,
        )
    )

    if failed:
        output.print_md('### Failed views:')
        for view_name, err in failed:
            output.print_md('- {} | {}'.format(view_name, err))

except Exception as err:
    output.print_md('## Error: {}'.format(str(err)))
    script.exit()
