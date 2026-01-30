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
)
from pyrevit import revit, script
from System.Windows.Forms import (
    Form,
    Button,
    DialogResult,
    TextBox,
    CheckedListBox,
    CheckBox,
    FormStartPosition,
)
from System.Collections.Generic import List
import re


# Button info
# ===================================================
__title__ = "Isolate by Fab Notes"
__doc__ = """
Select fabrication duct by Fabrication Notes parameter
"""

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

# Class
# =====================================================================


class EnhancedParamForm(Form):
    def __init__(self, param_groups):
        Form.__init__(self)
        self.Text = "Select Ducts by Fabrication Notes"
        self.Width = 600
        self.Height = 500
        self.StartPosition = FormStartPosition.CenterScreen
        self.param_groups = param_groups

        # Build flat list of values
        self.all_values = []
        for value in sorted(param_groups.keys(), key=natural_sort_key):
            count = len(param_groups[value])
            self.all_values.append((value, count))

        # Search box
        self.search_box = TextBox()
        self.search_box.Top = 10
        self.search_box.Left = 10
        self.search_box.Width = 560
        self.search_box.PlaceholderText = "Search..."
        self.search_box.TextChanged += self._filter_list
        self.Controls.Add(self.search_box)

        # CheckedListBox for flat list - optimized
        self.checked_list = CheckedListBox()
        self.checked_list.Top = 40
        self.checked_list.Left = 10
        self.checked_list.Width = 560
        self.checked_list.Height = 380
        self.checked_list.CheckOnClick = True  # Single click to check
        self.Controls.Add(self.checked_list)

        # Suspend layout for faster loading
        self.checked_list.BeginUpdate()
        self._build_list()
        self.checked_list.EndUpdate()

        # Select All and Itemize button
        btn_itemize = Button()
        btn_itemize.Text = "Select Checked"
        btn_itemize.Top = 430
        btn_itemize.Left = 10
        btn_itemize.Width = 150
        btn_itemize.DialogResult = DialogResult.Yes
        self.Controls.Add(btn_itemize)
        self.AcceptButton = btn_itemize

        # Fab Exported checkbox
        self.chk_fab_exported = CheckBox()
        self.chk_fab_exported.Text = "Fab Exported"
        self.chk_fab_exported.Top = 435
        self.chk_fab_exported.Left = 180
        self.chk_fab_exported.Width = 200
        self.Controls.Add(self.chk_fab_exported)

    def _on_itemize_click(self, sender, args):
        pass

    def _build_list(self, search_filter=None):
        self.checked_list.BeginUpdate()
        self.checked_list.Items.Clear()

        for value, count in self.all_values:
            if search_filter and search_filter not in str(value).lower():
                continue
            display_text = "{} ({} parts)".format(value, count)
            self.checked_list.Items.Add(display_text)

        self.checked_list.EndUpdate()

    def _filter_list(self, sender, args):
        search = sender.Text.lower()
        self._build_list(search if search else None)

    def get_checked_ducts(self):
        """Returns list of duct elements from all checked items"""
        ducts = set()
        for i in range(self.checked_list.CheckedItems.Count):
            item_text = str(self.checked_list.CheckedItems[i])
            value = item_text.rsplit(" (", 1)[0]
            if value in self.param_groups:
                for duct in self.param_groups[value]:
                    ducts.add(duct)
        return list(ducts)

    def require_fab_exported(self):
        return bool(self.chk_fab_exported.Checked)

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

    # Build parameter -> value -> elements map (only for Fabrication Notes)
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

        for p in list(d.Parameters):
            if p.Definition.Name != "Fabrication Notes":
                continue
            pval = get_param_value(p)
            if pval is None or pval == "":
                pval = "(blank)"
            if pval not in param_groups:
                param_groups[pval] = []
            param_groups[pval].append(d)
            break

    if not param_groups:
        script.exit()

    # Show form and get selection
    form = EnhancedParamForm(param_groups)
    if form.ShowDialog() != DialogResult.Yes:
        script.exit()

    duct_run = form.get_checked_ducts()
    if not duct_run:
        script.exit()

    # Optional filter: Fab Exported must be True
    if form.require_fab_exported():
        filtered = []
        for d in duct_run:
            p = d.LookupParameter("Fab Exported")
            if is_param_true(p):
                filtered.append(d)
        duct_run = filtered
        if not duct_run:
            script.exit()

    # Select ducts in Revit
    duct_ids = List[ElementId]()
    for d in duct_run:
        duct_ids.Add(d.Id)
    uidoc.Selection.SetElementIds(duct_ids)

    # Isolate selected ducts in the active view
    try:
        t = Transaction(doc, "Temporary Isolate Ducts")
        t.Start()
        view.IsolateElementsTemporary(duct_ids)
        t.Commit()
    except Exception:
        try:
            t.RollBack()
        except Exception:
            pass

except Exception as e:
    pass
