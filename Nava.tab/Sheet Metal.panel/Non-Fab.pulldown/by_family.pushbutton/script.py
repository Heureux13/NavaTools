# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from pyrevit import revit, script
from System.Windows.Forms import Form, Label, ComboBox, Button, DialogResult, ComboBoxStyle
from System.Collections.Generic import List
from revit_duct import RevitDuct
from revit_output import print_disclaimer
import clr
import re
clr.AddReference("System.Windows.Forms")


# Button info
# ===================================================
__title__ = "By Family"
__doc__ = """
Group and select all ducts by family that are NOT fab parts.
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


class SystemSelectorForm(object):
    def __init__(self, runs):
        self.form = Form()
        self.form.Text = "Select Duct Family"
        self.form.Width = 500
        self.form.Height = 180
        self.selected_family = None

        # Info label
        message = Label()
        message.Text = "Select family to select all its ducts:"
        message.Top = 20
        message.Left = 20
        message.Width = 350
        message.Height = 25
        self.form.Controls.Add(message)

        # System dropdown
        self.drop_down = ComboBox()
        self.drop_down.Top = 60
        self.drop_down.Left = 20
        self.drop_down.Width = 450
        self.drop_down.DropDownStyle = ComboBoxStyle.DropDownList

        # Sorts and creates the dropdown menu
        sorted_runs = sorted(runs.keys(), key=natural_sort_key)
        for run_name in sorted_runs:
            duct_count = len(runs[run_name])
            self.drop_down.Items.Add(
                "{} ({} duct parts)".format(
                    run_name,
                    duct_count
                )
            )

        if self.drop_down.Items.Count > 0:
            self.drop_down.SelectedIndex = 0

        self.form.Controls.Add(self.drop_down)

        # OK button
        btn_ok = Button()
        btn_ok.Text = "Select"
        btn_ok.Top = 100
        btn_ok.Left = 20
        btn_ok.Width = 80
        btn_ok.DialogResult = DialogResult.OK
        self.form.Controls.Add(btn_ok)
        self.form.AcceptButton = btn_ok

    def ShowDialog(self):
        return self.form.ShowDialog()

# Helpers
# ========================================================================


def natural_sort_key(s):
    # Sort runs with natural/numeric sorting
    return [
        int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)
    ]

# Main Code
# ==================================================


try:
    # Collect all non-fabricated ducts
    collector_0 = FilteredElementCollector(doc, view.Id)
    all_straights = collector_0.OfCategory(BuiltInCategory.OST_DuctCurves)\
        .WhereElementIsNotElementType()\
        .ToElements()

    # Collects ll non-fabricated fittins
    collector_1 = FilteredElementCollector(doc, view.Id)
    all_fittings = collector_1.OfCategory(BuiltInCategory.OST_DuctFitting)\
        .WhereElementIsNotElementType()\
        .ToElements()

    # Combines both list into one
    all_duct = list(all_straights) + list(all_fittings)

    # Checksto see if our combined list has values
    if not all_duct:
        TaskDialog.Show("No Ducts", "No ducts found in current view.")
        script.exit()

    # Group by Family
    duct_runs = {}
    for d in all_duct:
        try:
            # Get family name using RevitDuct class
            duct_obj = RevitDuct(doc, view, d)
            family_name = duct_obj.family if duct_obj.family else "Unknown"

            # Addes keys and values to dictionary
            if family_name not in duct_runs:
                duct_runs[family_name] = []
            duct_runs[family_name].append(d)

        # Handles erros from missing/bad data or parameters
        except Exception as e:
            output.print_md("Error grouping element: {}".format(str(e)))
            continue

    # Checks to see if list is empty before continuing
    if not duct_runs:
        output.print_md("Error, could not group ducts by system.")

    # Show window with dropdwn
    form = SystemSelectorForm(duct_runs)
    result = form.ShowDialog()

    # Stop script if user does not select OK
    if result != DialogResult.OK or not form.drop_down.SelectedItem:
        script.exit()

    # Convert our user selection to string, split it, and get our Family Name
    selected_text = str(form.drop_down.SelectedItem)
    family_name = selected_text.rsplit(" (", 1)[0]

    # Check to see if our family name exist
    if family_name not in duct_runs:
        TaskDialog.Show(
            "Family not found",
            "The selected family was not found in the grouped ducts."
        )
        script.exit()

    # Select all ducts in duct run
    duct_run = duct_runs[family_name]
    duct_ids = List[ElementId]([d.Id for d in duct_run])
    uidoc.Selection.SetElementIds(duct_ids)

    # final printout with links to duct
    for i, d in enumerate(duct_run, start=1):
        duct_obj = RevitDuct(doc, view, d)
        family_name = duct_obj.family if duct_obj.family else "Unknown"
        output.print_md(
            "### No: {:03} | ID: {} | Family: {}".format(
                i,
                output.linkify(d.Id),
                family_name
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
