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


# Button info
# ===================================================
__title__ = "MEP Ref Levels"
__doc__ = """
Collect all MEP elements in current view and print reference level differences
"""

# Variables
# ==================================================
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
view = revit.active_view
output = script.get_output()

# Helpers
# ========================================================================


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
        if level_id and level_id.IntegerValue > 0:
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
        if level_id and level_id.IntegerValue > 0:
            level = doc.GetElement(level_id)
            if level and hasattr(level, "Name"):
                return level.Name
    except Exception:
        pass

    return "<No Reference Level>"


def get_mep_category_ids():
    category_names = [
        "OST_DuctCurves",
        "OST_DuctFitting",
        "OST_DuctAccessory",
        "OST_FlexDuctCurves",
        "OST_FabricationDuctwork",
        "OST_PipeCurves",
        "OST_PipeFitting",
        "OST_PipeAccessory",
        "OST_FlexPipeCurves",
        "OST_FabricationPipework",
        "OST_CableTray",
        "OST_CableTrayFitting",
        "OST_Conduit",
        "OST_ConduitFitting",
        "OST_MechanicalEquipment",
        "OST_PlumbingFixtures",
        "OST_Sprinklers",
        "OST_AirTerminals",
        "OST_ElectricalEquipment",
        "OST_ElectricalFixtures",
        "OST_LightingFixtures",
        "OST_LightingDevices",
        "OST_DataDevices",
        "OST_FireAlarmDevices",
        "OST_CommunicationDevices",
        "OST_SecurityDevices",
        "OST_NurseCallDevices",
    ]

    category_ids = set()
    for name in category_names:
        bic = getattr(BuiltInCategory, name, None)
        if bic is not None:
            category_ids.add(int(bic))
    return category_ids


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
        if category.Id.IntegerValue in mep_category_ids:
            mep_elements.append(element)

    if not mep_elements:
        output.print_md("## No MEP elements found in current view")
        script.exit()

    level_groups = {}
    for element in mep_elements:
        level_name = get_reference_level_name(element)
        if level_name not in level_groups:
            level_groups[level_name] = []
        level_groups[level_name].append(element)

    selected_ids = List[ElementId]()
    for element in mep_elements:
        selected_ids.Add(element.Id)
    uidoc.Selection.SetElementIds(selected_ids)

    output.print_md("## Selected {} MEP elements in current view".format(len(mep_elements)))
    output.print_md("## Found {} different reference levels".format(len(level_groups)))
    output.print_md("---")

    sorted_levels = sorted(level_groups.keys(), key=lambda x: str(x).lower())
    for level_name in sorted_levels:
        level_elements = level_groups[level_name]
        output.print_md("### {} ({})".format(level_name, len(level_elements)))
        for element in sorted(level_elements, key=lambda x: x.Id.IntegerValue):
            category_name = element.Category.Name if element.Category else "<No Category>"
            output.print_md(
                "- ID: {} | Category: {}".format(
                    output.linkify(element.Id),
                    category_name
                )
            )

    element_ids = [element.Id for element in mep_elements]
    output.print_md(
        "# Total elements {}, {}".format(
            len(mep_elements),
            output.linkify(element_ids)
        )
    )

except Exception as e:
    output.print_md("## Error: {}".format(str(e)))
