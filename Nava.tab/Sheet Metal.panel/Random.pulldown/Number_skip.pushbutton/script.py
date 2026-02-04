# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from pyrevit import revit, forms
from Autodesk.Revit.DB import Transaction

# Button display information
# =================================================
__title__ = "Number Skip"
__doc__ = """
Sets parameters to 'skip' for selected elements
"""

# Parameters to set to 'skip'
parameters_to_skip = {
    "item number",
}

# Code
# ==================================================
uidoc = __revit__.ActiveUIDocument
doc = revit.doc

# Get selected elements
selected_ids = uidoc.Selection.GetElementIds()
if not selected_ids:
    forms.alert("Please select one or more elements", exitscript=True)

# Normalize parameter names for comparison
normalized_params_to_skip = {p.strip().lower() for p in parameters_to_skip}

t = Transaction(doc, "Set Parameters to Skip")
t.Start()
try:
    for elem_id in selected_ids:
        elem = doc.GetElement(elem_id)
        if elem is None:
            continue

        # Iterate through all element parameters
        for param in elem.Parameters:
            try:
                param_name = param.Definition.Name
                normalized_name = param_name.strip().lower()

                # Check if this parameter should be set to skip
                if normalized_name not in normalized_params_to_skip:
                    continue

                if param.IsReadOnly:
                    continue

                # Set the parameter value
                param.Set("skip")
            except Exception:
                pass

    t.Commit()
except Exception as e:
    t.RollBack()
    raise
