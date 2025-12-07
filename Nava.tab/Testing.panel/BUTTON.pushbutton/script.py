# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

import clr
clr.AddReference("System.Windows.Forms")

from System.Collections.Generic import List
from System.Windows.Forms import (
    Form,
    Label,
    ComboBox,
    Button,
    DialogResult,
    ComboBoxStyle
)
from pyrevit import revit, script
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.DB import *
from revit_duct import RevitDuct


# Button info
# ===================================================
__title__ = "Select Duct System"
__doc__ = """
Group and select all ducts by HVAC system.
"""

# Variables
# ==================================================
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
view = revit.active_view
output = script.get_output()

# Main Code
# ==================================================

try:
    # Collect all fabrication parts (not design ducts)
    all_fabs = RevitDuct.all(doc, view)
    all_fabs = FilteredElementCollector(doc, view.Id)\
        .OfCategory(BuiltInCategory.OST_DuctTerminal)\
        .WhereElementIsNotElementType()\
        .ToElements()

    if not all_fabs:
        TaskDialog.Show("No Fabrication Parts",
                        "No fabrication parts found in current view.")
        script.exit()

    # Group by System Name AND Fabrication Service
    duct_runs = {}
    for d in all_fabs:
        try:
            system_param = d.LookupParameter("System Name")
            system_name = system_param.AsString(
            ) if system_param and system_param.AsString() else "No System"

            fab_serv = d.LookupParameter("Fabrication Service")
            fab_service = fab_serv.AsString() if fab_serv and fab_serv.AsString() else system_name

            # Create combined key - if service is empty, just use system name
            if fab_service and fab_service != "No System":
                group_name = "{} - {}".format(system_name, fab_service)
            else:
                group_name = system_name

            # Debug output
            output.print_md("Element {}: System='{}', Service='{}'".format(
                d.Id, system_name, fab_service))

            if group_name not in duct_runs:
                duct_runs[group_name] = []
            duct_runs[group_name].append(d)
        except Exception as e:
            output.print_md("Error grouping element: {}".format(str(e)))
            continue

    if not duct_runs:
        TaskDialog.Show("Error", "Could not group ducts by system.")
        script.exit()

    # UI Form for system selection
    class SystemSelectorForm(Form):
        def __init__(self, runs):
            self.Text = "Select Duct System"
            self.Width = 500
            self.Height = 180

            # Text in box
            message = Label()
            message.Text = "Select system to select all its ducts:"
            message.Top = 20
            message.Left = 20
            message.Width = 350
            message.Height = 25
            self.Controls.Add(message)

            # System dropdown
            self.drop_down = ComboBox()
            self.drop_down.Top = 60
            self.drop_down.Left = 20
            self.drop_down.Width = 350
            self.drop_down.DropDownStyle = ComboBoxStyle.DropDownList

            # Sort runs with natural/numeric sorting
            def natural_sort_key(s):
                import re
                return [int(text) if text.isdigit() else text.lower()
                        for text in re.split(r'(\d+)', s)]

            sorted_runs = sorted(runs.keys(), key=natural_sort_key)
            for run_name in sorted_runs:
                duct_count = len(runs[run_name])
                self.drop_down.Items.Add(
                    "{} ({} parts)".format(run_name, duct_count))
            if self.drop_down.Items.Count > 0:
                self.drop_down.SelectedIndex = 0
            self.Controls.Add(self.drop_down)

            # OK button
            btn_ok = Button()
            btn_ok.Text = "Select"
            btn_ok.Top = 100
            btn_ok.Left = 20
            btn_ok.Width = 80
            btn_ok.DialogResult = DialogResult.OK
            self.Controls.Add(btn_ok)
            self.AcceptButton = btn_ok

    # Show form
    form = SystemSelectorForm(duct_runs)
    result = form.ShowDialog()

    if result != DialogResult.OK or not form.drop_down.SelectedItem:
        script.exit()

    # Parse selected system name (remove duct count from display)
    selected_text = str(form.drop_down.SelectedItem)
    system_name = selected_text.rsplit(" (", 1)[0]  # Remove " (X ducts)" part

    if system_name not in duct_runs:
        TaskDialog.Show("Error", "System not found.")
        script.exit()

    # Select all ducts in this system
    ducts = duct_runs[system_name]
    duct_ids = List[ElementId]([d.Id for d in ducts])
    uidoc.Selection.SetElementIds(duct_ids)

    # Report
    # output.print_md("**Selected {} ducts from system '{}'**".format(len(ducts), system_name))

except Exception as e:
    TaskDialog.Show("Error", "Script failed: {}".format(str(e)))
    import traceback
    output.print_md("```\n{}\n```".format(traceback.format_exc()))
