# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, VisibleInViewFilter, ElementId
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
__title__ = "Fab"
__doc__ = """
Selects all fabrication duct, and filters them down by parameters to select
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


class ParamSelectorForm(Form):
    def __init__(self, param_groups):
        Form.__init__(self)
        self.Text = "Select Ducts by Parameter"
        self.Width = 520
        self.Height = 220
        self.param_groups = param_groups

        # Parameter label
        lbl_param = Label()
        lbl_param.Text = "Choose parameter:"
        lbl_param.Top = 20
        lbl_param.Left = 20
        lbl_param.Width = 180
        lbl_param.Height = 20
        self.Controls.Add(lbl_param)

        # Parameter dropdown
        self.param_drop = ComboBox()
        self.param_drop.Top = 45
        self.param_drop.Left = 20
        self.param_drop.Width = 460
        self.param_drop.DropDownStyle = ComboBoxStyle.DropDownList
        for pname in sorted(param_groups.keys(), key=natural_sort_key):
            self.param_drop.Items.Add(pname)
        if self.param_drop.Items.Count > 0:
            self.param_drop.SelectedIndex = 0
        self.param_drop.SelectedIndexChanged += self._on_param_changed
        self.Controls.Add(self.param_drop)

        # Value label
        lbl_val = Label()
        lbl_val.Text = "Choose value:"
        lbl_val.Top = 80
        lbl_val.Left = 20
        lbl_val.Width = 180
        lbl_val.Height = 20
        self.Controls.Add(lbl_val)

        # Value dropdown
        self.value_drop = ComboBox()
        self.value_drop.Top = 105
        self.value_drop.Left = 20
        self.value_drop.Width = 460
        self.value_drop.DropDownStyle = ComboBoxStyle.DropDownList
        self.Controls.Add(self.value_drop)

        # OK button
        btn_ok = Button()
        btn_ok.Text = "Select"
        btn_ok.Top = 140
        btn_ok.Left = 20
        btn_ok.Width = 80
        btn_ok.DialogResult = DialogResult.OK
        self.Controls.Add(btn_ok)
        self.AcceptButton = btn_ok

        # init second dropdown
        self._populate_values()

    def _on_param_changed(self, sender, args):
        self._populate_values()

    def _populate_values(self):
        self.value_drop.Items.Clear()
        pname = self.param_drop.SelectedItem
        if not pname or pname not in self.param_groups:
            return
        values = self.param_groups[pname]
        for val in sorted(values.keys(), key=natural_sort_key):
            self.value_drop.Items.Add(
                "{} ({} parts)".format(val, len(values[val])))
        if self.value_drop.Items.Count > 0:
            self.value_drop.SelectedIndex = 0

# Helpers
# ========================================================================


def natural_sort_key(s):
    # Sort runs with natural/numeric sorting
    return [
        int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)
    ]


def get_param_value(param):
    try:
        if param.StorageType == 0:  # None
            return None
        # Prefer string if present
        if param.AsString():
            return param.AsString()
        # Next, use value string (formatted)
        if param.AsValueString():
            return param.AsValueString()
        # Finally, raw numeric
        if param.StorageType == 1:  # Double
            return param.AsDouble()
        if param.StorageType == 2:  # Integer
            return param.AsInteger()
        if param.StorageType == 3:  # ElementId
            return param.AsElementId().IntegerValue
    except Exception:
        return None
    return None


# Main Code
# ==================================================
try:
    # Collect only fabrication ductwork (visible in view)
    vis = VisibleInViewFilter(doc, view.Id)

    fab_duct = (FilteredElementCollector(doc, view.Id)
                .OfCategory(BuiltInCategory.OST_FabricationDuctwork)
                .WherePasses(vis)
                .WhereElementIsNotElementType()
                .ToElements())

    # Combines list into one (only fab ductwork available)
    all_duct = list(fab_duct)

    if not all_duct:
        TaskDialog.Show("No Ducts", "No ducts found in current view.")
        script.exit()

    # Build parameter -> value -> elements map
    param_groups = {}
    for d in all_duct:
        try:
            for p in list(d.Parameters):
                pname = p.Definition.Name
                pval = get_param_value(p)
                if pval is None or pval == "":
                    continue
                if pname not in param_groups:
                    param_groups[pname] = {}
                if pval not in param_groups[pname]:
                    param_groups[pname][pval] = []
                param_groups[pname][pval].append(d)
        except Exception as e:
            output.print_md("Error reading parameters: {}".format(str(e)))
            continue

    if not param_groups:
        TaskDialog.Show(
            "No Parameters",
            "No parameter data found on ducts in view."
        )
        script.exit()

    # Show window with parameter/value dropdowns
    form = ParamSelectorForm(param_groups)
    result = form.ShowDialog()

    # Stop script if user does not select OK
    if result != DialogResult.OK or not form.param_drop.SelectedItem or not form.value_drop.SelectedItem:
        script.exit()

    selected_param = str(form.param_drop.SelectedItem)
    selected_val = str(form.value_drop.SelectedItem).rsplit(" (", 1)[0]

    if selected_param not in param_groups or selected_val not in param_groups[selected_param]:
        TaskDialog.Show(
            "Not found", "The selection was not found in the grouped ducts.")
        script.exit()

    duct_run = param_groups[selected_param][selected_val]
    duct_ids = List[ElementId]()
    for d in duct_run:
        duct_ids.Add(d.Id)
    uidoc.Selection.SetElementIds(duct_ids)

    # final printout with links to duct
    for i, d in enumerate(duct_run, start=1):
        duct_obj = RevitDuct(doc, view, d)
        family_name = duct_obj.family if duct_obj.family else "Unknown"
        output.print_md(
            "### No: {} | ID: {} | Family: {}".format(
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
