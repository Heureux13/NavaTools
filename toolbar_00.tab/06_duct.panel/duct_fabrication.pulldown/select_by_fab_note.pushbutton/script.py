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
    FabricationPart,
)
from Autodesk.Revit.DB.Fabrication import FabricationSaveJobOptions
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent
from pyrevit import revit, script, forms
from System.Windows.Forms import (
    Form,
    Button,
    Label,
    TextBox,
    TreeView,
    TreeNode,
    CheckBox,
    FormStartPosition,
    AnchorStyles,
    SaveFileDialog,
    FolderBrowserDialog,
    DialogResult,
    MessageBox,
    MessageBoxButtons,
    MessageBoxIcon,
)
from System.Collections.Generic import List, HashSet
from System.Drawing import Size
from config.parameters_registry import PYT_DATE_EXPORT
import re
import os
import datetime


# Button info
# ===================================================
__title__ = "Select by Fab Note"
__doc__ = """
Select by Fab Notes with Item Number."""

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


class SelectDuctsEventHandler(IExternalEventHandler):
    """Runs Selection.SetElementIds() inside a valid Revit API context.

    WinForms button clicks execute outside Revit's API context once the
    form is shown modelessly, so the actual Selection call must be raised
    through an ExternalEvent instead of being called directly.
    """

    def __init__(self):
        self.duct_ids = None

    def Execute(self, uiapp):
        try:
            if self.duct_ids is not None:
                uiapp.ActiveUIDocument.Selection.SetElementIds(self.duct_ids)
        except Exception as ex:
            print("Select Ducts Event Handler failed: {}".format(ex))

    def GetName(self):
        return "Select Ducts Event Handler"


class ExportFabricationJobEventHandler(IExternalEventHandler):
    """Exports fabrication ducts to .maj file(s) inside a valid Revit API
    context, using FabricationPart.SaveAsFabricationJob().

    Supports a single export (duct_ids/file_path) or a batch export of
    several jobs (self.jobs), one .maj file per job.
    """

    def __init__(self):
        self.duct_ids = None
        self.file_path = None
        self.jobs = None  # list of dicts: {"label", "duct_ids", "file_path"}
        self.folders_created = 0

    def _save_one(self, doc, duct_ids, file_path):
        id_set = HashSet[ElementId]()
        for eid in duct_ids:
            id_set.Add(eid)

        # True = add holes for taps on straight sections (matches
        # the default behavior of Revit's "Export Job File" command).
        options = FabricationSaveJobOptions(True)
        exported_ids = FabricationPart.SaveAsFabricationJob(
            doc, id_set, file_path, options)

        self._stamp_export_date(doc, exported_ids)

        return exported_ids.Count if exported_ids else 0

    def _stamp_export_date(self, doc, exported_ids):
        """Write the current date/time (e.g. 2026.09.17_1440) to the
        PYT_DATE_EXPORT parameter of every successfully exported duct."""
        if not exported_ids or exported_ids.Count == 0:
            return

        stamp = datetime.datetime.now().strftime("%Y.%m.%d_%H%M")

        t = Transaction(doc, "Set PYT_DATE_EXPORT")
        t.Start()
        try:
            for eid in exported_ids:
                elem = doc.GetElement(eid)
                if not elem:
                    continue
                param = elem.LookupParameter(PYT_DATE_EXPORT)
                if param and not param.IsReadOnly:
                    param.Set(stamp)
            t.Commit()
        except Exception:
            t.RollBack()
            raise

    def Execute(self, uiapp):
        doc = uiapp.ActiveUIDocument.Document

        if self.jobs:
            jobs = self.jobs
            folders_created = self.folders_created
            self.jobs = None
            self.folders_created = 0

            files_created = 0
            failures = []
            for job in jobs:
                try:
                    self._save_one(doc, job["duct_ids"], job["file_path"])
                    files_created += 1
                except Exception as ex:
                    failures.append("{}: FAILED ({})".format(job["label"], ex))

            summary_lines = [
                "Files created: {}".format(files_created),
                "Folders created: {}".format(folders_created),
            ]
            if failures:
                summary_lines.append("")
                summary_lines.append("Failures:")
                summary_lines.extend(failures)

            MessageBox.Show(
                "\n".join(summary_lines),
                "Batch Export to MAJ",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)
            return

        try:
            if not self.duct_ids or not self.file_path:
                return

            count = self._save_one(doc, self.duct_ids, self.file_path)
            MessageBox.Show(
                "Exported {} fabrication part(s) to:\n{}".format(
                    count, self.file_path),
                "Export to MAJ",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)
        except Exception as ex:
            MessageBox.Show(
                "Export failed:\n{}".format(ex),
                "Export to MAJ",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)

    def GetName(self):
        return "Export Fabrication Job Event Handler"


class EnhancedParamForm(Form):
    def __init__(self, param_groups, uidoc, select_event=None, select_handler=None,
                 export_event=None, export_handler=None):
        Form.__init__(self)
        self.Text = "Select Ducts by Fabrication Notes"
        self.Width = 700
        self.Height = 600
        self.MinimumSize = Size(500, 400)
        self.StartPosition = FormStartPosition.CenterScreen
        self.param_groups = param_groups
        self.uidoc = uidoc
        self.select_event = select_event
        self.select_handler = select_handler
        self.export_event = export_event
        self.export_handler = export_handler

        # Build hierarchical structure: base_name -> [variants]
        self.hierarchy = self._build_hierarchy(param_groups)

        _MARGIN = 12
        _BTN_H = 30
        _NAME_H = 24
        _BTN_TOP = self.ClientSize.Height - _MARGIN - _BTN_H
        _EXPORT_BTN_TOP = _BTN_TOP - 6 - _BTN_H
        _NAME_TOP = _EXPORT_BTN_TOP - _MARGIN - _NAME_H
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
        self.tree_view.Height = _NAME_TOP - _tree_top - _MARGIN
        self.tree_view.CheckBoxes = True
        self.tree_view.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.tree_view.AfterCheck += self._on_node_checked
        self.Controls.Add(self.tree_view)

        # Build tree structure
        self._build_tree()

        # First Selected name box - shows the name of the first checked
        # group so it can be easily copied to the clipboard.
        _name_label_w = 100
        self.first_selected_label = Label()
        self.first_selected_label.Text = "First Selected:"
        self.first_selected_label.Top = _NAME_TOP + 4
        self.first_selected_label.Left = _MARGIN
        self.first_selected_label.Width = _name_label_w
        self.first_selected_label.Height = _NAME_H
        self.first_selected_label.Anchor = AnchorStyles.Bottom | AnchorStyles.Left
        self.Controls.Add(self.first_selected_label)

        self.first_selected_box = TextBox()
        self.first_selected_box.Top = _NAME_TOP
        self.first_selected_box.Left = _MARGIN + _name_label_w
        self.first_selected_box.Width = _inner_w - _name_label_w
        self.first_selected_box.Height = _NAME_H
        self.first_selected_box.ReadOnly = True
        self.first_selected_box.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.Controls.Add(self.first_selected_box)

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

        # Deselect All button
        btn_deselect = Button()
        btn_deselect.Text = "Deselect All"
        btn_deselect.Top = _BTN_TOP
        btn_deselect.Left = _MARGIN + 120 + 8 + 150 + 8
        btn_deselect.Width = 120
        btn_deselect.Height = _BTN_H
        btn_deselect.Anchor = AnchorStyles.Bottom | AnchorStyles.Left
        btn_deselect.Click += self._on_deselect_all
        self.Controls.Add(btn_deselect)

        # Close button
        btn_close = Button()
        btn_close.Text = "Close"
        btn_close.Top = _BTN_TOP
        btn_close.Left = _MARGIN + 120 + 8 + 150 + 8 + 120 + 8
        btn_close.Width = 100
        btn_close.Height = _BTN_H
        btn_close.Anchor = AnchorStyles.Bottom | AnchorStyles.Left
        btn_close.Click += self._on_close
        self.Controls.Add(btn_close)

        # Export to MAJ button (single file, everything checked)
        btn_export = Button()
        btn_export.Text = "Export to MAJ"
        btn_export.Top = _EXPORT_BTN_TOP
        btn_export.Left = _MARGIN
        btn_export.Width = 150
        btn_export.Height = _BTN_H
        btn_export.Anchor = AnchorStyles.Bottom | AnchorStyles.Left
        btn_export.Click += self._on_export_checked
        self.Controls.Add(btn_export)

        # Batch Export button (one .maj file per checked group)
        btn_batch_export = Button()
        btn_batch_export.Text = "Batch Export by Group"
        btn_batch_export.Top = _EXPORT_BTN_TOP
        btn_batch_export.Left = _MARGIN + 150 + 8
        btn_batch_export.Width = 180
        btn_batch_export.Height = _BTN_H
        btn_batch_export.Anchor = AnchorStyles.Bottom | AnchorStyles.Left
        btn_batch_export.Click += self._on_batch_export_checked
        self.Controls.Add(btn_batch_export)

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
        self._update_first_selected_box()

    def _on_select_checked(self, sender, args):
        """Select checked ducts in Revit without closing the form."""
        ducts = self.get_checked_ducts()
        if not ducts:
            return
        duct_ids = List[ElementId]()
        for d in ducts:
            duct_ids.Add(d.Id)

        if self.select_handler is not None and self.select_event is not None:
            # Modeless form: route the Selection call through an
            # ExternalEvent so it runs inside a valid Revit API context.
            self.select_handler.duct_ids = duct_ids
            self.select_event.Raise()
        else:
            self.uidoc.Selection.SetElementIds(duct_ids)

    def _on_export_checked(self, sender, args):
        """Export checked ducts to a .maj fabrication job file."""
        ducts = self.get_checked_ducts()
        if not ducts:
            MessageBox.Show(
                "No ducts are checked to export.",
                "Export to MAJ",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning)
            return

        if self.export_handler is None or self.export_event is None:
            MessageBox.Show(
                "Export is not available.",
                "Export to MAJ",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
            return

        dialog = SaveFileDialog()
        dialog.Title = "Export Selected Ducts to MAJ"
        dialog.Filter = "Fabrication Job (*.maj)|*.maj"
        dialog.DefaultExt = "maj"
        dialog.FileName = "export.maj"

        if dialog.ShowDialog() != DialogResult.OK:
            return

        duct_ids = List[ElementId]()
        for d in ducts:
            duct_ids.Add(d.Id)

        # Route the export through an ExternalEvent so the actual
        # FabricationPart.SaveAsFabricationJob() call runs inside a
        # valid Revit API context.
        self.export_handler.duct_ids = duct_ids
        self.export_handler.file_path = dialog.FileName
        self.export_event.Raise()

    def _on_batch_export_checked(self, sender, args):
        """Export each checked group to its own .maj file in a chosen folder."""
        groups = self.get_checked_groups()
        if not groups:
            MessageBox.Show(
                "No groups are checked to export.",
                "Batch Export to MAJ",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning)
            return

        if self.export_handler is None or self.export_event is None:
            MessageBox.Show(
                "Export is not available.",
                "Batch Export to MAJ",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
            return

        folder_dialog = FolderBrowserDialog()
        folder_dialog.Description = "Select a folder to export {} .maj file(s) to".format(len(groups))

        if folder_dialog.ShowDialog() != DialogResult.OK:
            return

        jobs = []
        folders_created = 0
        for variant, ducts in groups:
            duct_ids = List[ElementId]()
            for d in ducts:
                duct_ids.Add(d.Id)

            # Build nested folders: Equipment / System / Section, e.g.
            # "AHU-201 RA (A2)" -> AHU-201\RA\A2\AHU-201 RA (A2).maj
            equipment, system_abbrev, section = parse_variant_path(variant)
            if equipment and system_abbrev and section:
                dest_folder = os.path.join(
                    folder_dialog.SelectedPath,
                    sanitize_filename(equipment),
                    sanitize_filename(system_abbrev),
                    sanitize_filename(section))
            else:
                dest_folder = folder_dialog.SelectedPath

            if not os.path.isdir(dest_folder):
                os.makedirs(dest_folder)
                folders_created += 1

            file_name = sanitize_filename(variant) + ".maj"
            file_path = os.path.join(dest_folder, file_name)
            jobs.append({
                "label": variant,
                "duct_ids": duct_ids,
                "file_path": file_path,
            })

        # Route the batch export through an ExternalEvent so all the
        # FabricationPart.SaveAsFabricationJob() calls run inside a
        # valid Revit API context.
        self.export_handler.jobs = jobs
        self.export_handler.folders_created = folders_created
        self.export_event.Raise()

    def _on_deselect_all(self, sender, args):
        """Uncheck all nodes and clear the current Revit selection."""
        self.tree_view.AfterCheck -= self._on_node_checked
        self.tree_view.AfterCheck -= self._on_node_checked

        for parent_node in self.tree_view.Nodes:
            if parent_node.Tag and parent_node.Tag[0] == "parent":
                parent_node.Checked = False
                for child_node in parent_node.Nodes:
                    child_node.Checked = False

        self.tree_view.AfterCheck += self._on_node_checked
        self._update_first_selected_box()

        empty_ids = List[ElementId]()
        if self.select_handler is not None and self.select_event is not None:
            # Modeless form: route the Selection call through an
            # ExternalEvent so it runs inside a valid Revit API context.
            self.select_handler.duct_ids = empty_ids
            self.select_event.Raise()
        else:
            self.uidoc.Selection.SetElementIds(empty_ids)

    def _on_close(self, sender, args):
        """Close the modeless form."""
        self.Close()

    def _on_node_checked(self, sender, args):
        """When a parent is checked/unchecked, check/uncheck all children"""
        self.tree_view.AfterCheck -= self._on_node_checked

        node = args.Node

        # If parent node is checked/unchecked, apply to all children
        if node.Tag and node.Tag[0] == "parent":
            for child_node in node.Nodes:
                child_node.Checked = node.Checked

        self.tree_view.AfterCheck += self._on_node_checked
        self._update_first_selected_box()

    def _update_first_selected_box(self):
        """Show the name of the first checked group in the read-only box."""
        checked_values = self.get_checked_values()
        if checked_values:
            self.first_selected_box.Text = checked_values[0]
        else:
            self.first_selected_box.Text = ""

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

    def get_checked_groups(self):
        """Returns [(variant_name, [ducts]), ...] for each checked child
        node, keeping each variant's ducts as its own separate group
        (used for batch export, one file per group)."""
        groups = []

        for parent_node in self.tree_view.Nodes:
            if parent_node.Tag and parent_node.Tag[0] == "parent":
                for child_node in parent_node.Nodes:
                    if child_node.Checked and child_node.Tag and child_node.Tag[0] == "child":
                        variant = child_node.Tag[1]
                        ducts = self.param_groups.get(variant, [])
                        if ducts:
                            groups.append((variant, ducts))

        return groups

# Helpers
# ========================================================================


def natural_sort_key(s):
    # Sort runs with natural/numeric sorting
    return [
        int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)
    ]


def sanitize_filename(name):
    """Strip characters that are invalid in Windows file names."""
    cleaned = re.sub(r'[\\/:*?"<>|]', "_", name).strip()
    return cleaned or "export"


# Matches variant names like "AHU-201 RA (A2)" ->
#   equipment="AHU-201", system_abbrev="RA", section="A2"
_VARIANT_PATTERN = re.compile(r'^(.+?)\s+(\S+)\s*\(([^)]*)\)\s*$')


def parse_variant_path(variant):
    """Split a Fabrication Notes variant like "AHU-201 RA (A2)" into
    (equipment, system_abbrev, section) for building nested export
    folders. Returns (None, None, None) if it doesn't match the
    expected pattern.
    """
    match = _VARIANT_PATTERN.match(variant.strip())
    if not match:
        return None, None, None
    equipment, system_abbrev, section = match.groups()
    equipment = equipment.strip()
    section = section.strip()
    if not equipment or not system_abbrev or not section:
        return None, None, None
    return equipment, system_abbrev, section


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

    # Create the external event handlers up front so the modeless form can
    # safely call Revit API selection/export methods from a button click.
    select_handler = SelectDuctsEventHandler()
    select_event = ExternalEvent.Create(select_handler)

    export_handler = ExportFabricationJobEventHandler()
    export_event = ExternalEvent.Create(export_handler)

    # Show form modelessly — this keeps Revit responsive (you can still
    # click around in Revit) while the form stays open. Selection changes
    # and MAJ export triggered from the form run through the ExternalEvents
    # above.
    form = EnhancedParamForm(
        param_groups, uidoc, select_event, select_handler,
        export_event, export_handler)
    form.Show()

except Exception as e:
    output.print_md("**Error:** {}".format(str(e)))
    import traceback
    output.print_md("```\n{}\n```".format(traceback.format_exc()))
