# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from pyrevit import revit, script
from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    ElementParameterFilter,
    FilterDoubleRule,
    FilterNumericLess,
    FilteredElementCollector,
    ParameterValueProvider,
    Wall,
)
from Autodesk.Revit.UI import TaskDialog

# Button info
# ==================================================
__title__ = "Compare To C#"
__author__ = "Jose Francisco Nava Perez"
__doc__ = """
Python version of the C# sample: counts all walls and the walls
shorter than 12 feet using an ElementParameterFilter."""

# Variables
# ==================================================
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
output = script.get_output()

# Main Code
# ==================================================

# Collect all walls
walls = (FilteredElementCollector(doc)
         .OfClass(Wall)
         .WhereElementIsNotElementType()
         .ToElements())

TaskDialog.Show(doc.Title, "We have {} walls in the model".format(len(walls)))

# Construct a filter
parameter_id = ElementId(BuiltInParameter.WALL_USER_HEIGHT_PARAM)
provider = ParameterValueProvider(parameter_id)
rule = FilterNumericLess()
passes_rule = FilterDoubleRule(provider, rule, 12.0, 0.1)
param_filter = ElementParameterFilter(passes_rule)

# Collect all walls lower than 12 feet
walls_filtered = (FilteredElementCollector(doc)
                  .OfCategory(BuiltInCategory.OST_Walls)
                  .WhereElementIsNotElementType()
                  .WherePasses(param_filter)
                  .ToElements())

TaskDialog.Show(
    doc.Title,
    "We have {} walls less than 12 ft in the model".format(len(walls_filtered))
)
