# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, ElementId, CategoryType
from pyrevit import revit, script
from System.Collections.Generic import List
from System.Windows.Forms import Form, Label, Button, DialogResult, CheckedListBox
import clr

clr.AddReference("System.Windows.Forms")


# Button info
# ===================================================
__title__ = "Check Ref Level"
__doc__ = """
Collect MEP elements in current view, show a popup to pick reference levels, and select matching elements.
"""

# Variables
# ==================================================
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
view = revit.active_view
output = script.get_output()

# Helpers
# ========================================================================


def element_id_value(eid):
    if eid is None:
        return None
    return eid.IntegerValue if hasattr(eid, "IntegerValue") else eid.Value


def lookup_parameter_case_insensitive(element, param_name):
    """Case-insensitive parameter lookup"""
    param_name_lower = param_name.strip().lower()
    for param in element.Parameters:
        if param.Definition.Name.strip().lower() == param_name_lower:
            return param
    return None


def get_level_name_from_param(param):
    if not param:
        return None

    try:
        level_id = param.AsElementId()
        if level_id and element_id_value(level_id) > 0:
            level = doc.GetElement(level_id)
            if level and hasattr(level, "Name"):
                return level.Name
    except Exception:
        pass

    try:
        level_name = param.AsValueString() or param.AsString()
        if level_name:
            return str(level_name).strip()
    except Exception:
        pass

    return None


def get_reference_level_name(element):
    candidate_params = [
        "Reference Level",
        "Level",
        "Start Level",
        "Base Level",
    ]

    for param_name in candidate_params:
        param = lookup_parameter_case_insensitive(element, param_name)
        level_name = get_level_name_from_param(param)
        if level_name:
            return level_name

    try:
        level_id = element.LevelId
        if level_id and element_id_value(level_id) > 0:
            level = doc.GetElement(level_id)
            if level and hasattr(level, "Name"):
                return level.Name
    except Exception:
        pass

    return "<No Reference Level>"


def get_mep_category_ids():
    category_names = [
        "OST_DuctCurves",
        "OST_FabricationDuctwork",
        "OST_PipeCurves",
        "OST_FabricationPipework",
    ]

    category_ids = set()
    for name in category_names:
        bic = getattr(BuiltInCategory, name, None)
        if bic is not None:
            category_ids.add(int(bic))
    return category_ids


class LevelPickerForm(Form):
    def __init__(self, level_groups):
        Form.__init__(self)
        self.Text = "Pick Reference Levels"
        self.Width = 520
        self.Height = 620
        self.level_groups = level_groups
        self.display_to_level = {}

        title = Label()
        title.Text = "Select one or more reference levels"
        title.Left = 20
        title.Top = 15
        title.Width = 460
        self.Controls.Add(title)

        self.level_list = CheckedListBox()
        self.level_list.Left = 20
        self.level_list.Top = 45
        self.level_list.Width = 460
        self.level_list.Height = 480
        self.level_list.CheckOnClick = True
        self.Controls.Add(self.level_list)

        sorted_levels = sorted(level_groups.keys(),
                               key=lambda x: str(x).lower())
        for level_name in sorted_levels:
            count = len(level_groups[level_name])
            display = "{} ({})".format(level_name, count)
            idx = self.level_list.Items.Add(display)
            self.level_list.SetItemChecked(idx, True)
            self.display_to_level[display] = level_name

        btn_ok = Button()
        btn_ok.Text = "Select"
        btn_ok.Left = 20
        btn_ok.Top = 540
        btn_ok.Width = 120
        btn_ok.DialogResult = DialogResult.OK
        self.Controls.Add(btn_ok)
        self.AcceptButton = btn_ok

        btn_cancel = Button()
        btn_cancel.Text = "Cancel"
        btn_cancel.Left = 150
        btn_cancel.Top = 540
        btn_cancel.Width = 120
        btn_cancel.DialogResult = DialogResult.Cancel
        self.Controls.Add(btn_cancel)
        self.CancelButton = btn_cancel

    def get_selected_levels(self):
        selected = []
        for item in self.level_list.CheckedItems:
            display = str(item)
            level_name = self.display_to_level.get(display)
            if level_name:
                selected.append(level_name)
        return selected


# Main Code
# ==================================================
try:
    mep_category_ids = get_mep_category_ids()

    all_elements = (FilteredElementCollector(doc, view.Id)
                    .WhereElementIsNotElementType()
                    .ToElements())

    mep_elements = []
    for element in all_elements:
        category = element.Category
        if not category:
            continue
        if category.CategoryType != CategoryType.Model:
            continue
        if element_id_value(category.Id) in mep_category_ids:
            mep_elements.append(element)

    if not mep_elements:
        output.print_md("## No MEP elements found in current view.")
        script.exit()

    level_groups = {}
    for element in mep_elements:
        level_name = get_reference_level_name(element)
        if level_name not in level_groups:
            level_groups[level_name] = []
        level_groups[level_name].append(element)

    form = LevelPickerForm(level_groups)
    result = form.ShowDialog()

    if result != DialogResult.OK:
        script.exit()

    selected_levels = form.get_selected_levels()
    if not selected_levels:
        output.print_md("## Please select at least one reference level.")
        script.exit()

    selected_elements = []
    for level_name in selected_levels:
        selected_elements.extend(level_groups.get(level_name, []))

    if not selected_elements:
        output.print_md("## No elements found for selected reference levels.")
        script.exit()

    selected_ids = List[ElementId]()
    for element in selected_elements:
        selected_ids.Add(element.Id)
    uidoc.Selection.SetElementIds(selected_ids)

    output.print_md("## Reference Level Filter")
    output.print_md(
        "- Selected levels: {} of {}".format(len(selected_levels), len(level_groups)))
    for level_name in sorted(selected_levels, key=lambda x: str(x).lower()):
        output.print_md("  - {} ({})".format(level_name,
                        len(level_groups.get(level_name, []))))

    element_ids = [element.Id for element in selected_elements]
    output.print_md("---")
    output.print_md(
        "# Total Elements: {}, {}".format(
            len(selected_elements),
            output.linkify(element_ids)
        )
    )

except Exception as e:
    output.print_md("## Error: {}".format(str(e)))
