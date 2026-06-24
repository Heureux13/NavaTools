# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import DB, forms, revit, script
from System.Collections.Generic import List

# Button info
# ======================================================================
__title__ = 'Unhide Fab Tags in Views'
__doc__ = '''
Open a menu of all MEP Fabrication Ductwork Tag types.
Pick tag types and views, then unhide matching hidden tags in those views.
'''

# Variables
# ======================================================================

doc = revit.doc
uidoc = revit.uidoc
active_view = doc.ActiveView
output = script.get_output()


class ListOption(object):
    def __init__(self, item, display_name):
        self.item = item
        self.display_name = display_name


def _elem_name(elem):
    try:
        return elem.Name
    except Exception:
        return str(elem.Id)


def _tag_type_display_name(tag_type):
    family_name = "Unknown Family"
    try:
        fam = tag_type.Family
        if fam:
            family_name = fam.Name
    except Exception:
        pass
    return "{} | {}".format(family_name, _elem_name(tag_type))


def _view_display_name(view):
    try:
        return "{} | {}".format(view.ViewType, view.Name)
    except Exception:
        return str(view.Id)


def get_tag_family_and_type(tag_type):
    family_name = "(Unknown Family)"
    type_name = "(Unknown Type)"

    try:
        fam = tag_type.Family
        if fam and fam.Name:
            family_name = fam.Name
    except Exception:
        pass

    try:
        if tag_type.Name:
            type_name = tag_type.Name
    except Exception:
        pass

    return family_name, type_name


def get_tag_instance_family_and_type(tag):
    family_name = "(Unknown Family)"
    type_name = "(Unknown Type)"

    try:
        tag_type = doc.GetElement(tag.GetTypeId())
    except Exception:
        tag_type = None

    if not tag_type:
        return family_name, type_name

    try:
        fam = getattr(tag_type, 'Family', None)
        if fam and fam.Name:
            family_name = fam.Name
    except Exception:
        pass

    try:
        if tag_type.Name:
            type_name = tag_type.Name
    except Exception:
        pass

    return family_name, type_name


def _matches_selected_tag(elem, selected_type_id_values, selected_keys):
    type_id_value = None
    try:
        type_id_value = _get_element_id_value(elem.GetTypeId())
    except Exception:
        type_id_value = None

    by_id = type_id_value in selected_type_id_values
    fam_name, typ_name = get_tag_instance_family_and_type(elem)
    by_name = (fam_name, typ_name) in selected_keys
    return by_id or by_name


def collect_fabrication_duct_tag_types(document):
    return sorted(
        DB.FilteredElementCollector(document)
        .OfClass(DB.FamilySymbol)
        .OfCategory(DB.BuiltInCategory.OST_FabricationDuctworkTags)
        .ToElements(),
        key=lambda t: _tag_type_display_name(t).lower()
    )


def _get_element_id_value(element_id):
    try:
        return element_id.Value
    except Exception:
        pass

    try:
        return element_id.IntegerValue
    except Exception:
        pass

    try:
        return int(element_id)
    except Exception:
        return None


def _is_valid_target_view(view):
    if not view:
        return False

    try:
        if view.IsTemplate:
            return False
    except Exception:
        return False

    try:
        invalid_view_types = {
            'DrawingSheet',
            'ProjectBrowser',
            'SystemBrowser',
            'Undefined',
            'Internal',
        }
        if str(view.ViewType) in invalid_view_types:
            return False
    except Exception:
        pass

    return True


def _get_target_views_from_selection_or_active():
    selected_ids = list(uidoc.Selection.GetElementIds())
    selected_views = {}

    for element_id in selected_ids:
        elem = doc.GetElement(element_id)
        if not elem:
            continue

        try:
            if isinstance(elem, DB.Viewport):
                placed_view = doc.GetElement(elem.ViewId)
                if _is_valid_target_view(placed_view):
                    view_key = _get_element_id_value(placed_view.Id)
                    selected_views[view_key] = placed_view
                continue
        except Exception:
            pass

        try:
            if isinstance(elem, DB.View) and _is_valid_target_view(elem):
                view_key = _get_element_id_value(elem.Id)
                selected_views[view_key] = elem
        except Exception:
            pass

    if selected_views:
        return list(selected_views.values())

    if _is_valid_target_view(active_view):
        return [active_view]

    return []


def _collect_valid_target_views(document):
    views = [
        view for view in DB.FilteredElementCollector(document).OfClass(DB.View)
        if _is_valid_target_view(view)
    ]
    return sorted(views, key=lambda v: _view_display_name(v).lower())


def _prompt_for_target_views(document):
    valid_views = _collect_valid_target_views(document)
    view_options = [
        ListOption(view, '{} - {}'.format(idx, _view_display_name(view)))
        for idx, view in enumerate(valid_views, 1)
    ]
    if not view_options:
        return []

    selected_options = forms.SelectFromList.show(
        view_options,
        name_attr='display_name',
        multiselect=True,
        title='Select Views to Unhide Tags In'
    )
    if not selected_options:
        return []

    return [opt.item for opt in selected_options]


def collect_matching_tags_in_view(document, view, selected_type_id_values, selected_keys):
    matching_ids = []
    total_view_tags = 0

    # Category-based collection is more robust than class-based collection
    # for hidden annotation elements in some Revit versions/models.
    try:
        all_tags = (
            DB.FilteredElementCollector(document)
            .OfCategory(DB.BuiltInCategory.OST_FabricationDuctworkTags)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        all_tags = []

    for tag in all_tags:
        try:
            owner_view_id = None
            try:
                owner_view_id = tag.OwnerViewId
            except Exception:
                owner_view_id = None

            if _get_element_id_value(owner_view_id) != _get_element_id_value(view.Id):
                continue

            total_view_tags += 1

            if not _matches_selected_tag(tag, selected_type_id_values, selected_keys):
                continue

            matching_ids.append(tag.Id)
        except Exception:
            continue

    return matching_ids, total_view_tags


def collect_matching_hidden_ids_in_view(document, view, selected_type_id_values, selected_keys):
    matching_ids = []

    try:
        hidden_ids = list(view.GetHiddenElementIds() or [])
    except Exception:
        hidden_ids = []

    hidden_tag_total = 0
    for hidden_id in hidden_ids:
        try:
            elem = document.GetElement(hidden_id)
            if not elem:
                continue

            is_tag = isinstance(elem, DB.IndependentTag)
            if not is_tag:
                try:
                    cat = elem.Category
                    is_tag = bool(
                        cat and
                        _get_element_id_value(cat.Id) == int(DB.BuiltInCategory.OST_FabricationDuctworkTags)
                    )
                except Exception:
                    is_tag = False

            if not is_tag:
                continue

            hidden_tag_total += 1
            if not _matches_selected_tag(elem, selected_type_id_values, selected_keys):
                continue

            matching_ids.append(hidden_id)
        except Exception:
            continue

    return matching_ids, len(hidden_ids), hidden_tag_total


def collect_matching_tags_documentwide(document, selected_type_id_values, selected_keys):
    matching_ids = []
    total_tags = 0

    try:
        all_tags = (
            DB.FilteredElementCollector(document)
            .OfClass(DB.IndependentTag)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    except Exception:
        all_tags = []

    for tag in all_tags:
        try:
            total_tags += 1
            if not _matches_selected_tag(tag, selected_type_id_values, selected_keys):
                continue
            matching_ids.append(tag.Id)
        except Exception:
            continue

    return matching_ids, total_tags


def collect_tag_type_options(tag_types):
    by_family = {}
    for tag_type in tag_types:
        family_name, type_name = get_tag_family_and_type(tag_type)
        if family_name not in by_family:
            by_family[family_name] = {}
        if type_name not in by_family[family_name]:
            by_family[family_name][type_name] = []
        by_family[family_name][type_name].append(tag_type)
    return by_family


def select_tag_families_and_types(by_family):
    """Hierarchical TreeView selection with search and select-all support."""
    from System.Windows.Forms import (
        Button,
        DialogResult,
        Form,
        FormStartPosition,
        Label,
        TextBox,
        TreeNode,
        TreeView,
    )

    class TagSelectionForm(Form):
        def __init__(self, options_by_family):
            Form.__init__(self)
            self.Text = "Select Tag Families and Types"
            self.Width = 580
            self.Height = 620
            self.StartPosition = FormStartPosition.CenterScreen
            self._all_nodes = []
            self._by_family = options_by_family

            label_search = Label()
            label_search.Text = "Search:"
            label_search.Top = 12
            label_search.Left = 10
            label_search.Width = 50
            self.Controls.Add(label_search)

            self.search_box = TextBox()
            self.search_box.Top = 10
            self.search_box.Left = 64
            self.search_box.Width = 490
            self.search_box.TextChanged += self._filter_tree
            self.Controls.Add(self.search_box)

            self.tree_view = TreeView()
            self.tree_view.Top = 40
            self.tree_view.Left = 10
            self.tree_view.Width = 544
            self.tree_view.Height = 490
            self.tree_view.CheckBoxes = True
            self.tree_view.AfterCheck += self._on_node_checked
            self.Controls.Add(self.tree_view)

            btn_all = Button()
            btn_all.Text = "Select All"
            btn_all.Top = 538
            btn_all.Left = 10
            btn_all.Width = 110
            btn_all.Click += self._select_all
            self.Controls.Add(btn_all)

            btn_ok = Button()
            btn_ok.Text = "Select Checked"
            btn_ok.Top = 538
            btn_ok.Left = 310
            btn_ok.Width = 120
            btn_ok.DialogResult = DialogResult.OK
            self.Controls.Add(btn_ok)
            self.AcceptButton = btn_ok

            btn_cancel = Button()
            btn_cancel.Text = "Cancel"
            btn_cancel.Top = 538
            btn_cancel.Left = 440
            btn_cancel.Width = 110
            btn_cancel.DialogResult = DialogResult.Cancel
            self.Controls.Add(btn_cancel)
            self.CancelButton = btn_cancel

            self._build_tree()

        def _build_tree(self):
            self.tree_view.Nodes.Clear()
            self._all_nodes = []
            for family_name in sorted(self._by_family.keys()):
                types_dict = self._by_family[family_name]
                total_count = sum(len(types_dict[t]) for t in types_dict)
                parent = TreeNode("{} ({})".format(family_name, total_count))
                parent.Tag = ("family", family_name)

                for type_name in sorted(types_dict.keys()):
                    count = len(types_dict[type_name])
                    child = TreeNode("{} ({})".format(type_name, count))
                    child.Tag = (family_name, type_name)
                    parent.Nodes.Add(child)

                self.tree_view.Nodes.Add(parent)
                self._all_nodes.append(parent)

        def _filter_tree(self, sender, args):
            text = (self.search_box.Text or "").lower()
            self.tree_view.Nodes.Clear()

            for parent in self._all_nodes:
                parent_matches = (not text) or (text in parent.Text.lower())
                matching_children = []
                for child in parent.Nodes:
                    if (not text) or (text in child.Text.lower()):
                        matching_children.append(child)

                if not parent_matches and not matching_children:
                    continue

                parent_copy = TreeNode(parent.Text)
                parent_copy.Tag = parent.Tag
                parent_copy.Checked = parent.Checked

                for child in matching_children:
                    child_copy = TreeNode(child.Text)
                    child_copy.Tag = child.Tag
                    child_copy.Checked = child.Checked
                    parent_copy.Nodes.Add(child_copy)

                self.tree_view.Nodes.Add(parent_copy)

        def _check_node_recursive(self, node, checked):
            node.Checked = checked
            for child in node.Nodes:
                self._check_node_recursive(child, checked)

        def _select_all(self, sender, args):
            self.tree_view.AfterCheck -= self._on_node_checked
            for node in self.tree_view.Nodes:
                self._check_node_recursive(node, True)
            self.tree_view.AfterCheck += self._on_node_checked

        def _on_node_checked(self, sender, args):
            self.tree_view.AfterCheck -= self._on_node_checked
            node = args.Node
            if node.Tag and node.Tag[0] == "family":
                for child in node.Nodes:
                    child.Checked = node.Checked
            self.tree_view.AfterCheck += self._on_node_checked

        def get_checked_keys(self):
            selected = set()

            def _walk(node):
                if node.Checked and node.Tag and node.Tag[0] != "family":
                    selected.add(node.Tag)
                for child in node.Nodes:
                    _walk(child)

            for root in self.tree_view.Nodes:
                _walk(root)
            return selected

    picker = TagSelectionForm(by_family)
    if picker.ShowDialog() == DialogResult.OK:
        return picker.get_checked_keys()
    return None


tag_types = collect_fabrication_duct_tag_types(doc)
if not tag_types:
    forms.alert('No MEP Fabrication Ductwork Tag types found in this project.', exitscript=True)

by_family = collect_tag_type_options(tag_types)
selected_keys = select_tag_families_and_types(by_family)
if not selected_keys:
    script.exit()

selected_type_id_values = set()
for family_name, type_name in selected_keys:
    for tag_type in by_family.get(family_name, {}).get(type_name, []):
        id_value = _get_element_id_value(tag_type.Id)
        if id_value is not None:
            selected_type_id_values.add(id_value)

if not selected_type_id_values:
    script.exit()

selected_views = _get_target_views_from_selection_or_active()
if not selected_views:
    selected_views = _prompt_for_target_views(doc)

if not selected_views:
    forms.alert('No valid target views found or selected.', exitscript=True)
    script.exit()

unhidden_count = 0
view_results = []

with revit.Transaction('Unhide fabrication duct tags in selected views'):
    documentwide_ids, doc_tag_total = collect_matching_tags_documentwide(
        doc, selected_type_id_values, selected_keys)

    for view in selected_views:
        # Ensure the category itself is visible, in case the tags are hidden
        # by category/annotation visibility rather than per-element hide.
        try:
            cat_id = DB.ElementId(DB.BuiltInCategory.OST_FabricationDuctworkTags)
            if view.CanCategoryBeHidden(cat_id):
                view.SetCategoryHidden(cat_id, False)
        except Exception:
            pass

        ids_to_try, total_view_tags = collect_matching_tags_in_view(
            doc, view, selected_type_id_values, selected_keys)
        hidden_ids_to_try, hidden_total, hidden_tag_total = collect_matching_hidden_ids_in_view(
            doc, view, selected_type_id_values, selected_keys)

        ids_to_try = list(set(list(ids_to_try) + list(hidden_ids_to_try) + list(documentwide_ids)))

        if not ids_to_try:
            view_results.append(
                (_view_display_name(view),
                 0,
                 'No matching tags found in this view (visible tags: {}, hidden ids: {}, hidden tags: {})'.format(
                    total_view_tags,
                    hidden_total,
                    hidden_tag_total)))
            continue

        success_count = 0
        failed_count = 0
        for tag_id in ids_to_try:
            try:
                view.UnhideElements(List[DB.ElementId]([tag_id]))
                success_count += 1
            except Exception:
                failed_count += 1

        unhidden_count += success_count
        if success_count:
            view_results.append((_view_display_name(view), success_count, None))
        else:
            view_results.append(
                (_view_display_name(view),
                 0,
                 'No tags were unhidden (matching tags: {}, failed attempts: {}, document tags scanned: {})'.format(
                    len(ids_to_try),
                    failed_count,
                    doc_tag_total)))

output.print_md('## Unhide Fabrication Ductwork Tags')
output.print_md('Views selected: {}'.format(len(selected_views)))
output.print_md('Tags unhidden: {}'.format(unhidden_count))
for view_name, count, error in view_results:
    if error:
        output.print_md('- {}: {} ({})'.format(view_name, count, error))
    else:
        output.print_md('- {}: {}'.format(view_name, count))
