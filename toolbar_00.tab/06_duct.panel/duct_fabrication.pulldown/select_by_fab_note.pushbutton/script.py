# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

import datetime
import os
import re
from config.parameters_registry import PYT_DATE_EXPORT
from System.Collections.Generic import List, HashSet
from System.Windows.Forms import (
    SaveFileDialog,
    FolderBrowserDialog,
    DialogResult,
    MessageBox,
    MessageBoxButtons,
    MessageBoxIcon,
)
from pyrevit import revit, script, forms
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent
from Autodesk.Revit.DB.Fabrication import FabricationSaveJobOptions
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    ElementId,
    VisibleInViewFilter,
    Transaction,
    FabricationPart,
)
from System.Windows.Controls import TreeViewItem, CheckBox
from System.Windows.Markup import XamlReader
import clr
clr.AddReference("WindowsBase")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("System.Windows.Forms")


# Button info
# ===================================================
__title__ = "Select by Fab Note"
__doc__ = """
Select by Fab Notes with Item Number in a modeless WPF window.
Search and check groups, select ducts, or export single and grouped MAJ files."""
__persistentengine__ = True
__fullframeengine__ = True

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

    Modeless button clicks execute outside Revit's API context once the
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


FAB_NOTES_XAML = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Select Ducts by Fabrication Notes"
        Width="740" Height="640" MinWidth="620" MinHeight="420"
        WindowStartupLocation="CenterScreen">
  <Grid Margin="12">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <DockPanel Margin="0,0,0,10">
      <TextBlock Text="Search:" VerticalAlignment="Center" Margin="0,0,10,0"/>
      <TextBox x:Name="search_box" Padding="5"/>
    </DockPanel>
    <TreeView x:Name="tree_view" Grid.Row="1"/>
    <TextBlock x:Name="summary_text" Grid.Row="2" Margin="0,8,0,8"/>
    <DockPanel Grid.Row="3" Margin="0,0,0,10">
      <TextBlock Text="First Selected:" VerticalAlignment="Center" Margin="0,0,10,0"/>
      <TextBox x:Name="first_selected_box" IsReadOnly="True" Padding="5"/>
    </DockPanel>
    <WrapPanel Grid.Row="4" Margin="0,0,0,6">
      <Button x:Name="export_button" Content="Export to MAJ" Padding="12,6" Margin="0,0,8,4"/>
      <Button x:Name="batch_export_button" Content="Batch Export by Group" Padding="12,6" Margin="0,0,8,4"/>
    </WrapPanel>
    <WrapPanel Grid.Row="5">
      <Button x:Name="all_button" Content="Select All" Padding="12,6" Margin="0,0,8,4"/>
      <Button x:Name="select_button" Content="Select Checked" Padding="12,6" Margin="0,0,8,4"/>
      <Button x:Name="deselect_button" Content="Deselect All" Padding="12,6" Margin="0,0,8,4"/>
      <Button x:Name="close_button" Content="Close" Padding="12,6" Margin="0,0,8,4"/>
    </WrapPanel>
  </Grid>
</Window>
"""


class EnhancedParamForm(object):
    def __init__(self, param_groups, uidoc, select_event=None, select_handler=None,
                 export_event=None, export_handler=None):
        self.window = XamlReader.Parse(FAB_NOTES_XAML)
        for name in (
                "search_box", "tree_view", "summary_text", "first_selected_box",
                "export_button", "batch_export_button", "all_button",
                "select_button", "deselect_button", "close_button"):
            control = self.window.FindName(name)
            if control is None:
                raise RuntimeError("WPF layout is missing '{}'.".format(name))
            setattr(self, name, control)
        self.param_groups = param_groups
        self.uidoc = uidoc
        self.select_event = select_event
        self.select_handler = select_handler
        self.export_event = export_event
        self.export_handler = export_handler

        self.hierarchy = self._build_hierarchy(param_groups)
        self.checked_variants = set()
        self.updating = False
        self.search_box.TextChanged += self._filter_tree
        self.all_button.Click += self._on_select_all
        self.select_button.Click += self._on_select_checked
        self.deselect_button.Click += self._on_deselect_all
        self.close_button.Click += self._on_close
        self.export_button.Click += self._on_export_checked
        self.batch_export_button.Click += self._on_batch_export_checked
        self._build_tree()

    def Show(self):
        self.window.Show()

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
        self.updating = True
        try:
            self.tree_view.Items.Clear()
            for base_name in sorted(self.hierarchy, key=natural_sort_key):
                variants = self.hierarchy[base_name]
                base_matches = not search_filter or search_filter in base_name.lower()
                visible = [
                    v for v in sorted(variants, key=natural_sort_key)
                    if base_matches or search_filter in v.lower()]
                if not visible:
                    continue
                parent = TreeViewItem()
                parent.Tag = visible
                parent.IsExpanded = bool(search_filter)
                checkbox = CheckBox()
                checkbox.Content = "{} ({} parts)".format(
                    base_name, sum(len(self.param_groups[v]) for v in visible))
                checkbox.Tag = ("parent", parent)
                parent.Header = checkbox
                for variant in visible:
                    child = TreeViewItem()
                    child.Tag = variant
                    child_box = CheckBox()
                    child_box.Content = "{} ({} parts)".format(
                        variant, len(self.param_groups[variant]))
                    child_box.Tag = ("child", parent, variant)
                    child_box.IsChecked = variant in self.checked_variants
                    child_box.Checked += self._on_node_checked
                    child_box.Unchecked += self._on_node_checked
                    child.Header = child_box
                    parent.Items.Add(child)
                self._update_parent(parent)
                checkbox.Checked += self._on_node_checked
                checkbox.Unchecked += self._on_node_checked
                self.tree_view.Items.Add(parent)
        finally:
            self.updating = False
        self._update_first_selected_box()

    def _update_parent(self, parent):
        states = [variant in self.checked_variants for variant in parent.Tag]
        parent.Header.IsChecked = True if all(states) else (None if any(states) else False)

    def _filter_tree(self, sender, args):
        """Filter tree based on search text"""
        search = sender.Text.lower()
        self._build_tree(search if search else None)

    def _on_select_all(self, sender, args):
        """Check all groups visible under the current search."""
        for parent in self.tree_view.Items:
            self.checked_variants.update(parent.Tag)
        self._filter_tree(self.search_box, None)

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
        self.checked_variants.clear()
        self._filter_tree(self.search_box, None)

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
        self.window.Close()

    def _on_node_checked(self, sender, args):
        args.Handled = True
        if self.updating:
            return
        self.updating = True
        try:
            parent = sender.Tag[1]
            variants = parent.Tag if sender.Tag[0] == "parent" else [sender.Tag[2]]
            for variant in variants:
                if sender.IsChecked:
                    self.checked_variants.add(variant)
                else:
                    self.checked_variants.discard(variant)
            for child in parent.Items:
                child.Header.IsChecked = child.Tag in self.checked_variants
            self._update_parent(parent)
        finally:
            self.updating = False
        self._update_first_selected_box()

    def _update_first_selected_box(self):
        """Show the name of the first checked group in the read-only box."""
        checked_values = self.get_checked_values()
        if checked_values:
            self.first_selected_box.Text = checked_values[0]
        else:
            self.first_selected_box.Text = ""
        self.summary_text.Text = "{} groups checked | {} parts".format(
            len(checked_values), len(self.get_checked_ducts()))

    def get_checked_ducts(self):
        """Returns list of duct elements from all checked nodes"""
        ducts = set()

        for variant in self.checked_variants:
            for duct in self.param_groups[variant]:
                ducts.add(duct)

        return list(ducts)

    def get_checked_values(self):
        """Returns selected Fabrication Notes values from checked child nodes"""
        return sorted(self.checked_variants, key=natural_sort_key)

    def get_checked_groups(self):
        """Returns [(variant_name, [ducts]), ...] for each checked child
        node, keeping each variant's ducts as its own separate group
        (used for batch export, one file per group)."""
        return [(variant, self.param_groups[variant])
                for variant in self.get_checked_values() if self.param_groups[variant]]

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
