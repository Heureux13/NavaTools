# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
import traceback
import clr

clr.AddReference("WindowsBase")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("System.Data")

from config.parameters_registry import *
from ducts.revit_xyz import RevitXYZ
from System.Windows import Thickness
from System.Windows.Markup import XamlReader
from System.Windows.Controls import ListBoxItem, Button
from System.Data import DataTable
from System.Collections.Generic import List
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent, ExternalEventRequest
from Autodesk.Revit.DB import ElementId, StorageType, BuiltInParameter
from pyrevit import revit, forms, script


# Button display information
# =================================================
__title__ = "Basic Info"
__doc__ = """
Inspect selected elements in a modeless WPF window.
Search parameters, refresh from the current selection, and select or zoom to IDs."""
__persistentengine__ = True
__fullframeengine__ = True

# Variables
# ==================================================
app = __revit__.Application             # type: Application
uidoc = __revit__.ActiveUIDocument        # type: UIDocument
doc = revit.doc                         # type: Document
view = revit.active_view
output = script.get_output()

show_parameters = {
    RVT_SIZE,
    RVT_ITEM_NUMBER,
    RVT_SPECIFICATION,
    RVT_INSULATION_SPECIFICATION,
    RVT_FABRICATION_SERVICE,
    RVT_MARK,
    RVT_TYPE_MARK,
    RVT_ANGLE,
    RVT_REFERENCE_LEVEL,
    RVT_UPPER_END_TOP_ELEVATION,
    RVT_MIDDLE_ELEVATION,
    RVT_LOWER_END_BOTTOM_ELEVATION,
}

show_parameters_compare = {param.lower().strip() for param in show_parameters}


def id_value(element_id):
    return element_id.Value if hasattr(element_id, "Value") else element_id.IntegerValue


def parameter_text(param):
    if param is None:
        return "N/A"
    value = param.AsString() if param.StorageType == StorageType.String else None
    if value is None:
        value = param.AsValueString()
    if value is None:
        if param.StorageType == StorageType.Double:
            value = param.AsDouble()
        elif param.StorageType == StorageType.Integer:
            value = param.AsInteger()
        elif param.StorageType == StorageType.ElementId:
            value = id_value(param.AsElementId())
    return str(value) if value is not None else "None"


def collect_info(document, element):
    info = {"id": element.Id, "parameters": [], "connected": [], "errors": []}
    element_type = document.GetElement(element.GetTypeId())
    for source, owner in (("Instance", element), ("Type", element_type)):
        if owner is None:
            continue
        for param in owner.Parameters:
            name = param.Definition.Name
            if name.lower().strip() not in show_parameters_compare:
                continue
            try:
                value = parameter_text(param)
            except Exception as ex:
                value = "Error: {}".format(ex)
                info["errors"].append("{} [{}]: {}".format(name, source, ex))
            info["parameters"].append((name, value, source))
    info["parameters"].sort(key=lambda row: (row[0].lower(), row[2]))

    try:
        info["family"] = parameter_text(
            element.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM))
        info["type"] = parameter_text(
            element.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM))
    except Exception as ex:
        info["family"] = info["type"] = "Unavailable"
        info["errors"].append("Family/type: {}".format(ex))

    try:
        xyz = RevitXYZ(element)
        start, end = xyz.curve_endpoints()
        if start is not None and end is not None:
            location = "Start ({:.3f}, {:.3f}, {:.3f})\nEnd ({:.3f}, {:.3f}, {:.3f})".format(
                start.X, start.Y, start.Z, end.X, end.Y, end.Z)
        else:
            origins = xyz.connector_origins()
            location = "\n".join(
                "Connector {}: ({:.3f}, {:.3f}, {:.3f})".format(i, p.X, p.Y, p.Z)
                for i, p in enumerate(origins or [])) or "Not available"
        info["location"] = location
    except Exception as ex:
        info["location"] = "Unavailable"
        info["errors"].append("Location: {}".format(ex))

    try:
        manager = getattr(element, "ConnectorManager", None)
        if manager is None:
            model = getattr(element, "MEPModel", None)
            manager = model.ConnectorManager if model is not None else None
        connected = {}
        if manager is not None:
            for connector in manager.Connectors:
                for reference in connector.AllRefs:
                    owner = reference.Owner
                    value = id_value(owner.Id)
                    if value == id_value(element.Id) or value in connected:
                        continue
                    connected[value] = (
                        owner.Id, parameter_text(owner.LookupParameter("Item Number")))
        info["connected"] = [connected[key] for key in sorted(connected)]
    except Exception as ex:
        info["errors"].append("Connected elements: {}".format(ex))
    return info


XAML = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Basic Duct Information" Width="1050" Height="700"
        MinWidth="750" MinHeight="450" WindowStartupLocation="CenterScreen">
  <Grid Margin="12">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/><RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <DockPanel Margin="0,0,0,10">
      <Button x:Name="refresh_button" Content="Refresh from Revit Selection"
              Padding="12,6" DockPanel.Dock="Left"/>
      <TextBlock Text="Select elements in Revit, then click Refresh."
                 VerticalAlignment="Center" Margin="12,0,0,0"/>
    </DockPanel>
    <Grid Grid.Row="1">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="230"/><ColumnDefinition Width="8"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      <ListBox x:Name="element_list"/>
      <GridSplitter Grid.Column="1" Width="5" HorizontalAlignment="Center"/>
      <Grid Grid.Column="2">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <StackPanel Margin="0,0,0,10">
          <TextBlock x:Name="element_heading" FontSize="18" FontWeight="Bold"/>
          <TextBlock x:Name="family_text" TextWrapping="Wrap" Margin="0,5"/>
          <Button x:Name="zoom_button" Content="Select and Zoom to Element"
                  HorizontalAlignment="Left" Padding="10,4"/>
        </StackPanel>
        <DockPanel Grid.Row="1" Margin="0,0,0,8">
          <TextBlock Text="Search parameters:" VerticalAlignment="Center" Margin="0,0,8,0"/>
          <TextBox x:Name="search_box" Padding="4"/>
        </DockPanel>
        <TabControl Grid.Row="2">
          <TabItem Header="Parameters">
            <DataGrid x:Name="parameter_grid" AutoGenerateColumns="False"
                      IsReadOnly="True" CanUserAddRows="False" SelectionMode="Extended"
                      ClipboardCopyMode="IncludeHeader">
              <DataGrid.Columns>
                <DataGridTextColumn Header="Parameter" Binding="{Binding Name}" Width="2*"/>
                <DataGridTextColumn Header="Value" Binding="{Binding Value}" Width="2*"/>
                <DataGridTextColumn Header="Source" Binding="{Binding Source}" Width="*"/>
              </DataGrid.Columns>
            </DataGrid>
          </TabItem>
          <TabItem Header="Location">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
              <StackPanel Margin="10">
                <TextBlock Text="Coordinates in Revit internal feet" FontWeight="Bold"/>
                <TextBox x:Name="location_text" IsReadOnly="True" TextWrapping="Wrap"
                         BorderThickness="0" Margin="0,10,0,0"/>
              </StackPanel>
            </ScrollViewer>
          </TabItem>
          <TabItem Header="Connected Elements">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
              <StackPanel x:Name="connected_panel" Margin="10"/>
            </ScrollViewer>
          </TabItem>
          <TabItem Header="Read Errors">
            <TextBox x:Name="errors_text" IsReadOnly="True" TextWrapping="Wrap"
                     VerticalScrollBarVisibility="Auto" Margin="10"/>
          </TabItem>
        </TabControl>
      </Grid>
    </Grid>
    <TextBlock x:Name="status_text" Grid.Row="2" TextWrapping="Wrap" Margin="0,10,0,0"/>
  </Grid>
</Window>
"""


class InfoEventHandler(IExternalEventHandler):
    def __init__(self, document):
        self.document = document
        self.window = None
        self.pending = False
        self.action = None
        self.element_id = None

    def Execute(self, uiapp):
        try:
            if self.window is None or not self.window.window.IsVisible:
                return
            active = uiapp.ActiveUIDocument
            if (active is None or not self.document.IsValidObject
                    or not active.Document.Equals(self.document)):
                raise RuntimeError("Activate the original document before using this window.")
            if self.action == "refresh":
                snapshots = []
                for eid in active.Selection.GetElementIds():
                    element = self.document.GetElement(eid)
                    if element is not None:
                        snapshots.append(collect_info(self.document, element))
                self.window.set_elements(snapshots)
            elif self.action == "zoom":
                if self.document.GetElement(self.element_id) is None:
                    raise RuntimeError("This element no longer exists. Refresh the information.")
                ids = List[ElementId]()
                ids.Add(self.element_id)
                active.Selection.SetElementIds(ids)
                active.ShowElements(ids)
                self.window.status_text.Text = "Selected element {}.".format(id_value(self.element_id))
        except Exception as ex:
            if self.window is not None:
                self.window.status_text.Text = "Error: {}".format(ex)
            print(traceback.format_exc())
        finally:
            self.pending = False
            self.action = None
            self.element_id = None

    def GetName(self):
        return "Basic Information Refresh and Selection"


class InfoWindow(object):
    def __init__(self, snapshots, handler, event):
        self.window = XamlReader.Parse(XAML)
        for name in (
                "refresh_button", "zoom_button", "element_list", "search_box",
                "element_heading", "family_text", "location_text", "errors_text",
                "connected_panel", "parameter_grid", "status_text"):
            control = self.window.FindName(name)
            if control is None:
                raise RuntimeError("WPF layout is missing control '{}'.".format(name))
            setattr(self, name, control)
        if self.window.Content is None:
            raise RuntimeError("WPF layout loaded without window content.")
        self.handler = handler
        self.event = event
        self.current = None
        self.table = None
        self.refresh_button.Click += self._refresh
        self.zoom_button.Click += self._zoom
        self.element_list.SelectionChanged += self._element_changed
        self.search_box.TextChanged += self._search_changed
        self.window.Closed += self._closed
        self.set_elements(snapshots)

    def set_elements(self, snapshots):
        self.element_list.Items.Clear()
        self.current = None
        self._show_current()
        for info in snapshots:
            item = ListBoxItem()
            item.Content = "ID: {}".format(id_value(info["id"]))
            item.Tag = info
            self.element_list.Items.Add(item)
        if snapshots:
            self.element_list.SelectedIndex = 0
        self.status_text.Text = "{} element(s) loaded. Data is a snapshot; use Refresh to update.".format(
            len(snapshots))

    def _element_changed(self, sender, args):
        item = self.element_list.SelectedItem
        self.current = item.Tag if item is not None else None
        self._show_current()

    def _show_current(self):
        info = self.current
        self.zoom_button.IsEnabled = info is not None
        self.element_heading.Text = "Element {}".format(id_value(info["id"])) if info else "No element selected"
        self.family_text.Text = "Family: {} | Type: {}".format(info["family"], info["type"]) if info else ""
        self.location_text.Text = info["location"] if info else ""
        self.errors_text.Text = "\n".join(info["errors"]) or "No read errors." if info else ""
        self.connected_panel.Children.Clear()
        if info:
            for eid, item_number in info["connected"]:
                button = Button()
                button.Content = "ID: {} | Item #: {} (select and zoom)".format(
                    id_value(eid), item_number)
                button.Tag = eid
                button.Margin = Thickness(0, 0, 0, 6)
                button.Padding = Thickness(8, 5, 8, 5)
                button.Click += self._connected_clicked
                self.connected_panel.Children.Add(button)
        self._fill_parameters()

    def _fill_parameters(self):
        table = DataTable()
        for column in ("Name", "Value", "Source"):
            table.Columns.Add(column)
        search = self.search_box.Text.strip().lower()
        for name, value, source in self.current["parameters"] if self.current else []:
            if not search or search in "{} {} {}".format(name, value, source).lower():
                table.Rows.Add(name, value, source)
        self.parameter_grid.ItemsSource = table.DefaultView
        previous = self.table
        self.table = table
        if previous is not None:
            previous.Dispose()

    def _search_changed(self, sender, args):
        self._fill_parameters()

    def _request(self, action, eid=None):
        if self.handler.pending:
            self.status_text.Text = "A Revit request is pending. Finish the current Revit command and try again."
            return
        try:
            self.handler.action = action
            self.handler.element_id = eid
            self.handler.pending = True
            result = self.event.Raise()
            if result != ExternalEventRequest.Accepted:
                raise RuntimeError("Revit did not accept the request: {}".format(result))
            self.status_text.Text = "Waiting for Revit..."
        except Exception as ex:
            self.handler.pending = False
            self.handler.action = None
            self.handler.element_id = None
            self.status_text.Text = "Error: {}".format(ex)
            print(traceback.format_exc())

    def _refresh(self, sender, args):
        self._request("refresh")

    def _zoom(self, sender, args):
        if self.current is not None:
            self._request("zoom", self.current["id"])

    def _connected_clicked(self, sender, args):
        self._request("zoom", sender.Tag)

    def _closed(self, sender, args):
        self.handler.window = None
        self.event.Dispose()
        self.parameter_grid.ItemsSource = None
        if self.table is not None:
            self.table.Dispose()


selected_ids = uidoc.Selection.GetElementIds()
if not selected_ids:
    forms.alert("Please select one or more elements.", exitscript=True)

try:
    snapshots = [collect_info(doc, doc.GetElement(eid)) for eid in selected_ids]
    info_handler = InfoEventHandler(doc)
    info_event = ExternalEvent.Create(info_handler)
    info_window = InfoWindow(snapshots, info_handler, info_event)
    info_handler.window = info_window
    info_window.window.Show()
except Exception as ex:
    forms.alert("Could not open Basic Information:\n{}".format(ex))
    output.print_md("```\n{}\n```".format(traceback.format_exc()))
