# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from constants.print_outputs import print_disclaimer
from pyrevit import revit, script
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    IndependentTag,
    Transaction,
    ElementId,
    BuiltInParameter,
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
__title__ = "Change Annotations"
__doc__ = """
Change one annotation type to another using a two-step tree picker.
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


def get_type_identity(elem_type):
    if elem_type is None:
        return "", ""

    family_name = ""
    type_name = ""

    try:
        fam = getattr(elem_type, "Family", None)
        family_name = (fam.Name if fam else "").strip()
    except Exception:
        family_name = ""

    if not family_name:
        family_name = (getattr(elem_type, "FamilyName", "") or "").strip()

    if not family_name:
        try:
            family_name = get_param_text(
                elem_type.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
            )
        except Exception:
            family_name = ""

    try:
        type_name = (getattr(elem_type, "Name", "") or "").strip()
    except Exception:
        type_name = ""

    if not type_name:
        try:
            type_name = get_param_text(
                elem_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            )
        except Exception:
            type_name = ""

    if not family_name:
        family_name = "(No Family Name)"
    if not type_name:
        type_name = "(No Type Name)"

    return family_name, type_name


def build_type_catalog(tag_types, counts_by_type_id=None):
    """Build category -> family -> [type rows] for the tree picker."""
    catalog = {}
    for tag_type in tag_types:
        if tag_type is None:
            continue

        type_id = tag_type.Id.IntegerValue

        try:
            cat_name = (tag_type.Category.Name or "").strip()
        except Exception:
            cat_name = ""
        if not cat_name:
            cat_name = "(No Category)"

        family_name, type_name = get_type_identity(tag_type)

        if cat_name not in catalog:
            catalog[cat_name] = {}
        if family_name not in catalog[cat_name]:
            catalog[cat_name][family_name] = []

        count = 0
        if counts_by_type_id:
            count = counts_by_type_id.get(type_id, 0)

        catalog[cat_name][family_name].append({
            "type_id": type_id,
            "category": cat_name,
            "family": family_name,
            "type": type_name,
            "count": count,
        })

    for cat_name in catalog.keys():
        for family_name in catalog[cat_name].keys():
            catalog[cat_name][family_name] = sorted(
                catalog[cat_name][family_name],
                key=lambda row: natural_sort_key(row["type"]),
            )

    return catalog


class AnnotationTypePicker(Form):
    def __init__(self, title, prompt, catalog, show_counts=False):
        Form.__init__(self)
        self.Text = title
        self.StartPosition = FormStartPosition.CenterScreen
        self.Width = 760
        self.Height = 760
        self.catalog = catalog
        self.show_counts = show_counts
        self.selected_type_id = None
        self.selected_label = ""

        self.lbl_prompt = Label()
        self.lbl_prompt.Text = prompt
        self.lbl_prompt.Top = 10
        self.lbl_prompt.Left = 20
        self.lbl_prompt.Width = 700
        self.lbl_prompt.Height = 42
        self.lbl_prompt.AutoSize = False
        self.Controls.Add(self.lbl_prompt)

        self.search_box = TextBox()
        self.search_box.Top = 55
        self.search_box.Left = 20
        self.search_box.Width = 700
        self.search_box.TextChanged += self._on_search_changed
        self.Controls.Add(self.search_box)

        self.tree_view = TreeView()
        self.tree_view.Top = 85
        self.tree_view.Left = 20
        self.tree_view.Width = 700
        self.tree_view.Height = 590
        self.tree_view.HideSelection = False
        self.tree_view.CheckBoxes = True
        self.tree_view.AfterCheck += self._on_after_check
        self.Controls.Add(self.tree_view)

        self.lbl_selected = Label()
        self.lbl_selected.Text = "Selected: none"
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

        for cat_name in sorted(self.catalog.keys(), key=natural_sort_key):
            families = self.catalog[cat_name]
            cat_node = TreeNode(cat_name)
            cat_node.Tag = ("category", cat_name)

            for family_name in sorted(families.keys(), key=natural_sort_key):
                types = families[family_name]
                fam_node = TreeNode(family_name)
                fam_node.Tag = ("family", cat_name, family_name)

                for row in types:
                    type_text = row["type"]
                    type_id = row["type_id"]
                    count = row["count"]

                    haystack = "{} {} {}".format(cat_name, family_name, type_text).lower()
                    if search and search not in haystack:
                        continue

                    if self.show_counts:
                        node_text = "{} ({} in view)".format(type_text, count)
                    else:
                        node_text = type_text

                    type_node = TreeNode(node_text)
                    type_node.Tag = ("type", type_id, cat_name, family_name, type_text)
                    fam_node.Nodes.Add(type_node)

                if fam_node.Nodes.Count > 0:
                    cat_node.Nodes.Add(fam_node)

            if cat_node.Nodes.Count > 0:
                self.tree_view.Nodes.Add(cat_node)

        # Keep tree collapsed by default so users opt into expanded branches.
        self.tree_view.CollapseAll()

    def _on_search_changed(self, sender, args):
        self._build_tree(sender.Text)

    def _on_after_check(self, sender, args):
        node = args.Node
        if node is None:
            return

        # Prevent recursive check events while we propagate parent checks.
        self.tree_view.AfterCheck -= self._on_after_check
        try:
            if node.Tag and node.Tag[0] in ("category", "family"):
                for child in node.Nodes:
                    child.Checked = node.Checked
        finally:
            self.tree_view.AfterCheck += self._on_after_check

        checked_types = self.get_checked_type_nodes()
        if len(checked_types) == 1:
            _, type_id, cat_name, family_name, type_name = checked_types[0].Tag
            self.selected_type_id = int(type_id)
            self.selected_label = "{} | {} | {}".format(cat_name, family_name, type_name)
            self.lbl_selected.Text = "Selected: {}".format(self.selected_label)
        elif len(checked_types) == 0:
            self.selected_type_id = None
            self.selected_label = ""
            self.lbl_selected.Text = "Selected: none"
        else:
            self.selected_type_id = None
            self.selected_label = ""
            self.lbl_selected.Text = "Selected: {} type nodes".format(len(checked_types))

    def get_checked_type_nodes(self):
        checked = []

        for cat_node in self.tree_view.Nodes:
            for fam_node in cat_node.Nodes:
                for type_node in fam_node.Nodes:
                    if type_node.Checked and type_node.Tag and type_node.Tag[0] == "type":
                        checked.append(type_node)

        return checked

    def _on_ok_clicked(self, sender, args):
        checked_types = self.get_checked_type_nodes()
        if len(checked_types) != 1:
            TaskDialog.Show("Select Type", "Check exactly one type node before clicking Select.")
            self.DialogResult = getattr(DialogResult, "None")
            return

        _, type_id, cat_name, family_name, type_name = checked_types[0].Tag
        self.selected_type_id = int(type_id)
        self.selected_label = "{} | {} | {}".format(cat_name, family_name, type_name)


def collect_view_independent_tags(doc_obj, view_obj):
    return list(
        FilteredElementCollector(doc_obj, view_obj.Id)
        .OfClass(IndependentTag)
        .WhereElementIsNotElementType()
        .ToElements()
    )


def collect_all_independent_tag_types(doc_obj):
    types = []
    for elem in FilteredElementCollector(doc_obj).WhereElementIsElementType().ToElements():
        if elem is None:
            continue
        try:
            if not elem.Category:
                continue
        except Exception:
            continue

        try:
            if elem.Category.CategoryType.ToString() != "Annotation":
                continue
        except Exception:
            continue

        # Keep to tag-like annotation families to avoid text/dimension types in this tool.
        cat_name = (elem.Category.Name or "").lower()
        if "tag" not in cat_name:
            continue
        types.append(elem)
    return types


def type_counts_for_tags(tags):
    counts = {}
    for tag in tags:
        if tag is None:
            continue
        try:
            tid = tag.GetTypeId()
            if tid is None:
                continue
            key = tid.IntegerValue
        except Exception:
            continue
        counts[key] = counts.get(key, 0) + 1
    return counts


def find_tags_by_type_id(tags, type_id_int):
    result = []
    for tag in tags:
        if tag is None:
            continue
        try:
            tid = tag.GetTypeId()
            if tid and tid.IntegerValue == type_id_int:
                result.append(tag)
        except Exception:
            continue
    return result


try:
    view_tags = collect_view_independent_tags(doc, view)
    if not view_tags:
        TaskDialog.Show("No Annotations", "No IndependentTag annotations found in this view.")
        script.exit()

    all_tag_types = collect_all_independent_tag_types(doc)
    if not all_tag_types:
        TaskDialog.Show("No Annotation Types", "No tag annotation types found in this project.")
        script.exit()

    counts = type_counts_for_tags(view_tags)
    source_type_ids = set(counts.keys())
    source_types = [t for t in all_tag_types if t.Id.IntegerValue in source_type_ids]

    if not source_types:
        TaskDialog.Show("No Source Types", "No matching annotation types were found for this view.")
        script.exit()

    source_catalog = build_type_catalog(source_types, counts_by_type_id=counts)
    src_form = AnnotationTypePicker(
        "Choose Annotation To Change",
        "Step 1: Select the source annotation type currently in this view.",
        source_catalog,
        show_counts=True,
    )
    if src_form.ShowDialog() != DialogResult.OK:
        script.exit()

    source_type_id = src_form.selected_type_id
    if source_type_id is None:
        script.exit()

    source_display = src_form.selected_label or "Type ID {}".format(source_type_id)

    target_catalog = build_type_catalog(all_tag_types, counts_by_type_id=None)
    tgt_form = AnnotationTypePicker(
        "Choose Replacement Annotation",
        "Step 2: Change FROM [{}]. Select the annotation type to swap TO (from all project tag types).".format(source_display),
        target_catalog,
        show_counts=False,
    )
    if tgt_form.ShowDialog() != DialogResult.OK:
        script.exit()

    target_type_id = tgt_form.selected_type_id
    if target_type_id is None:
        script.exit()

    if target_type_id == source_type_id:
        TaskDialog.Show("No Change", "Source and target types are the same.")
        script.exit()

    source_tags = find_tags_by_type_id(view_tags, source_type_id)
    if not source_tags:
        TaskDialog.Show("No Matching Tags", "No tags of the selected source type were found in this view.")
        script.exit()

    changed = 0
    skipped = 0
    changed_ids = []

    tx = Transaction(doc, "Change annotation type")
    tx.Start()
    try:
        target_eid = ElementId(int(target_type_id))

        for tag in source_tags:
            try:
                current_id = tag.GetTypeId()
                if current_id and current_id.IntegerValue == target_type_id:
                    skipped += 1
                    continue

                tag.ChangeTypeId(target_eid)
                changed += 1
                changed_ids.append(tag.Id)
            except Exception:
                skipped += 1

        tx.Commit()
    except Exception:
        tx.RollBack()
        raise

    if changed_ids:
        selected_ids = List[ElementId]()
        for ann_id in changed_ids:
            selected_ids.Add(ann_id)
        uidoc.Selection.SetElementIds(selected_ids)

    output.print_md("# Changed annotations in current view")
    output.print_md("- Source: {}".format(src_form.selected_label or source_type_id))
    output.print_md("- Target: {}".format(tgt_form.selected_label or target_type_id))
    output.print_md("- Changed: {}".format(changed))
    output.print_md("- Skipped (incompatible or failed): {}".format(skipped))

    if changed_ids and len(changed_ids) <= 500:
        output.print_md("## Changed Annotation IDs")
        for i, ann_id in enumerate(changed_ids, start=1):
            output.print_md("{}. {}".format(i, output.linkify(ann_id)))

    print_disclaimer(output)

except Exception as exc:
    TaskDialog.Show("Error", "Script failed: {}".format(str(exc)))
    import traceback
    output.print_md("```\n{}\n```".format(traceback.format_exc()))
