# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
import clr
clr.AddReference("WindowsBase")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("System.Data")

from System import Boolean, String
from System.Data import DataTable
from System.Windows.Markup import XamlReader
from System.Windows.Controls import DataGridEditingUnit
from pyrevit import revit, forms, script
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    RevitLinkInstance,
    ElementId,
    TemporaryViewMode,
)
from System.Collections.Generic import List
import sys
import json
import os

# Button info
# ===================================================
__title__ = "Hide Links"
__doc__ = """
Search visible Revit links, edit local nicknames, and temporarily hide checked links.
Nicknames do not rename the Revit links themselves.
Reset Temporary Hide/Isolate restores all temporarily hidden or isolated elements
in the active view, not just links. Permanent hiding is unchanged."""

# Variables
# ==================================================
doc = revit.doc
active_view = doc.ActiveView
output = script.get_output()

NICKNAME_KEY = 'NavaTools_LinkNicknames'


def get_element_id_value(element_id):
    """Return an integer ElementId value across Revit 2025 and 2026."""
    if element_id is None:
        return None
    try:
        return element_id.IntegerValue
    except AttributeError:
        try:
            return int(element_id.Value)
        except AttributeError:
            return None


def get_all_nicknames():
    try:
        data_file = script.get_document_data_file(doc.PathName or 'unsaved', NICKNAME_KEY)
        if os.path.exists(data_file):
            with open(data_file, 'r') as f:
                return json.load(f)
    except Exception as ex:
        forms.alert("Could not read link nicknames:\n{}".format(ex), exitscript=True)
    return {}


def save_all_nicknames(nicknames_dict):
    try:
        data_file = script.get_document_data_file(doc.PathName or 'unsaved', NICKNAME_KEY)
        with open(data_file, 'w') as f:
            json.dump(nicknames_dict, f)
    except Exception as ex:
        raise RuntimeError("Could not save link nicknames:\n{}".format(ex))


def get_link_nickname(link, nicknames_dict):
    # Check custom nicknames first
    link_id_value = get_element_id_value(link.Id)
    if link_id_value is None:
        return None
    link_id_str = str(link_id_value)
    if link_id_str in nicknames_dict:
        return nicknames_dict[link_id_str]
    return None


links_all = list(
    FilteredElementCollector(doc)
    .OfClass(RevitLinkInstance)
    .WhereElementIsNotElementType()
)

if not links_all and not active_view.IsTemporaryHideIsolateActive():
    output.print_md('**No Revit links found in this model.**')
    sys.exit(0)

# Filter to only visible links (collector with view ID gets visible elements)
links_in_view = list(
    FilteredElementCollector(doc, active_view.Id)
    .OfClass(RevitLinkInstance)
    .WhereElementIsNotElementType()
)

nicknames_dict = get_all_nicknames()

if not links_in_view and not active_view.IsTemporaryHideIsolateActive():
    forms.alert("No Revit links are visible in the active view.", exitscript=True)


LINKS_XAML = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Hide Links and Edit Nicknames" Width="900" Height="550"
        MinWidth="700" MinHeight="350" WindowStartupLocation="CenterScreen">
  <Grid Margin="12">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/><RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock Text="Check links to temporarily hide. Edit nicknames directly; these do not rename Revit links."
               TextWrapping="Wrap" Margin="0,0,0,10"/>
    <DockPanel Grid.Row="1" Margin="0,0,0,10">
      <TextBlock Text="Search:" VerticalAlignment="Center" Margin="0,0,8,0"/>
      <TextBox x:Name="search_box" Padding="4"/>
    </DockPanel>
    <DataGrid x:Name="links_grid" Grid.Row="2" AutoGenerateColumns="False"
              CanUserAddRows="False" CanUserDeleteRows="False"
              SelectionMode="Extended">
      <DataGrid.Columns>
        <DataGridCheckBoxColumn Header="Hide" Binding="{Binding Checked, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" Width="55"/>
        <DataGridTextColumn Header="Nickname" Binding="{Binding Nickname, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" Width="2*"/>
        <DataGridTextColumn Header="Revit Link" Binding="{Binding LinkName}" IsReadOnly="True" Width="3*"/>
        <DataGridTextColumn Header="Element ID" Binding="{Binding ElementId}" IsReadOnly="True" Width="110"/>
      </DataGrid.Columns>
    </DataGrid>
    <TextBlock x:Name="status_text" Grid.Row="3" Margin="0,8" TextWrapping="Wrap"/>
    <WrapPanel Grid.Row="4">
      <Button x:Name="all_button" Content="Check Visible" Padding="10,6" Margin="0,0,8,4"/>
      <Button x:Name="clear_button" Content="Uncheck All" Padding="10,6" Margin="0,0,8,4"/>
      <Button x:Name="save_button" Content="Save Nicknames" Padding="10,6" Margin="0,0,8,4"/>
      <Button x:Name="hide_button" Content="Hide Checked" Padding="10,6" Margin="0,0,8,4"/>
      <Button x:Name="reset_button" Content="Reset Temporary Hide/Isolate"
              ToolTip="Restores all temporarily hidden or isolated elements in this view, not just links."
              Padding="10,6" Margin="0,0,8,4"/>
      <Button x:Name="cancel_button" Content="Close" Padding="10,6" Margin="0,0,8,4"/>
    </WrapPanel>
  </Grid>
</Window>
"""


class LinkSelectionForm(object):
    def __init__(self, links, nicknames):
        self.window = XamlReader.Parse(LINKS_XAML)
        for name in ("search_box", "links_grid", "status_text", "all_button",
                     "clear_button", "save_button", "hide_button", "reset_button", "cancel_button"):
            control = self.window.FindName(name)
            if control is None:
                raise RuntimeError("WPF layout is missing '{}'.".format(name))
            setattr(self, name, control)
        self.nicknames = dict(nicknames)
        self.selected_ids = []
        self.reset_requested = False
        self.id_map = {str(get_element_id_value(link.Id)): link.Id for link in links}
        self.table = DataTable()
        self.table.Columns.Add("Checked", Boolean)
        for name in ("Nickname", "LinkName", "ElementId"):
            self.table.Columns.Add(name, String)
        for link in sorted(links, key=lambda item: item.Name.lower()):
            self.table.Rows.Add(
                False, get_link_nickname(link, nicknames) or "",
                link.Name, str(get_element_id_value(link.Id)))
        self.links_grid.ItemsSource = self.table.DefaultView
        self.search_box.TextChanged += self._search
        self.all_button.Click += self._check_visible
        self.clear_button.Click += self._uncheck_all
        self.save_button.Click += self._save
        self.hide_button.Click += self._hide
        self.reset_button.Click += self._reset
        self.reset_button.IsEnabled = active_view.IsTemporaryHideIsolateActive()
        self.cancel_button.Click += self._cancel
        self.status_text.Text = "{} visible link(s). Hide Checked also saves nickname edits.".format(
            len(links))

    def _commit(self):
        if (not self.links_grid.CommitEdit(DataGridEditingUnit.Cell, True)
                or not self.links_grid.CommitEdit(DataGridEditingUnit.Row, True)):
            raise RuntimeError("Finish editing the current cell before continuing.")

    def _search(self, sender, args):
        try:
            self._commit()
            text = sender.Text.replace("'", "''")
            text = "".join(
                {"[": "[[]", "]": "[]]", "%": "[%]", "*": "[*]"}.get(char, char)
                for char in text)
            self.table.DefaultView.RowFilter = (
                "Nickname LIKE '%{0}%' OR LinkName LIKE '%{0}%' OR ElementId LIKE '%{0}%'".format(text)
                if text else "")
        except Exception as ex:
            forms.alert("Could not filter links:\n{}".format(ex))

    def _check_visible(self, sender, args):
        try:
            self._commit()
            for row in self.table.DefaultView:
                row["Checked"] = True
        except Exception as ex:
            forms.alert("Could not check links:\n{}".format(ex))

    def _uncheck_all(self, sender, args):
        try:
            self._commit()
            for row in self.table.Rows:
                row["Checked"] = False
        except Exception as ex:
            forms.alert("Could not uncheck links:\n{}".format(ex))

    def _save_nicknames(self):
        self._commit()
        updated = dict(self.nicknames)
        for row in self.table.Rows:
            nickname = str(row["Nickname"]).strip()
            if nickname:
                updated[row["ElementId"]] = nickname
            else:
                updated.pop(row["ElementId"], None)
        save_all_nicknames(updated)
        self.nicknames = updated
        self.status_text.Text = "Nicknames saved. Blank nicknames remove the saved alias."

    def _save(self, sender, args):
        try:
            self._save_nicknames()
        except Exception as ex:
            forms.alert(str(ex))

    def _hide(self, sender, args):
        try:
            self._commit()
            selected = [
                self.id_map[row["ElementId"]] for row in self.table.Rows if row["Checked"]]
            if not selected:
                forms.alert("Check at least one link to hide.")
                return
            self._save_nicknames()
            self.selected_ids = selected
            self.window.DialogResult = True
        except Exception as ex:
            forms.alert("Could not prepare links to hide:\n{}".format(ex))

    def _cancel(self, sender, args):
        self.window.DialogResult = False

    def _reset(self, sender, args):
        self.reset_requested = True
        self.window.DialogResult = True


try:
    form = LinkSelectionForm(links_in_view, nicknames_dict)
    try:
        confirmed = form.window.ShowDialog()
    finally:
        form.links_grid.ItemsSource = None
        form.table.Dispose()
    if confirmed:
        if form.reset_requested:
            with revit.Transaction('Reset Temporary Hide/Isolate'):
                active_view.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate)
        else:
            ids_to_hide = List[ElementId]()
            for link_id in form.selected_ids:
                ids_to_hide.Add(link_id)
            with revit.Transaction('Hide Selected Links'):
                active_view.HideElementsTemporary(ids_to_hide)
except Exception as ex:
    forms.alert("Hide Links failed:\n{}".format(ex))
