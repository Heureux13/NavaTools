# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from pyrevit import revit, forms, script
from Autodesk.Revit.DB import *

# Button display information
# =================================================
__title__ = "Parameters"
__doc__ = """
Shows all available parameters for selected element.
"""

# Variables
# ==================================================
app = __revit__.Application             # type: Application
uidoc = __revit__.ActiveUIDocument        # type: UIDocument
doc = revit.doc                         # type: Document
view = revit.active_view
output = script.get_output()

# Main Code
# ==================================================
selected_ids = uidoc.Selection.GetElementIds()
if not selected_ids:
    forms.alert("please select one or more elements", exitscript=True)

output.print_md("# ELEMENT PARAMETERS")
output.print_md("---")

for elid in selected_ids:
    el = doc.GetElement(elid)
    output.print_md("###---- Parameters for Element ID {} ----".format(el.Id))

    # Collect instance parameters
    param_list = []
    for param in el.Parameters:
        try:
            name = param.Definition.Name
            value = param.AsString()
            if value is None:
                value = param.AsValueString()
            if value is None:
                if param.StorageType == StorageType.Double:
                    value = param.AsDouble()
                elif param.StorageType == StorageType.Integer:
                    value = param.AsInteger()
                elif param.StorageType == StorageType.ElementId:
                    value = param.AsElementId()
            param_list.append((name, value, "Instance"))
        except Exception as ex:
            param_list.append((name, "Error - {}".format(ex), "Instance"))

    # Collect type parameters
    elem_type = doc.GetElement(el.GetTypeId())
    if elem_type:
        for param in elem_type.Parameters:
            try:
                name = param.Definition.Name
                value = param.AsString()
                if value is None:
                    value = param.AsValueString()
                if value is None:
                    if param.StorageType == StorageType.Double:
                        value = param.AsDouble()
                    elif param.StorageType == StorageType.Integer:
                        value = param.AsInteger()
                    elif param.StorageType == StorageType.ElementId:
                        value = param.AsElementId()
                param_list.append((name, value, "Type"))
            except Exception as ex:
                param_list.append((name, "Error - {}".format(ex), "Type"))

    # Sort and print
    for name, value, param_type in sorted(param_list, key=lambda x: x[0].lower()):
        display_value = value if value is not None else "None"
        output.print_md("**{}** [{}]:     *{}*".format(name, param_type, display_value))
