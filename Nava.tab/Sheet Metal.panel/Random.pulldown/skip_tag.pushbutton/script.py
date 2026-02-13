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
__title__ = "Skip Tag"
__doc__ = """
Sets parameters to skip selected elements when tagging
"""

# Parameters to set to 'skip'
parameters_to_skip = {
    "_duct_tag_offset",
    '_duct_tag',
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

        # Set parameters only if they exist on the element
        for param_name in normalized_params_to_skip:
            try:
                param = elem.LookupParameter(param_name)
                if param is None or param.IsReadOnly:
                    continue

                param.Set("skip")
            except Exception:
                pass

    t.Commit()
except Exception as e:
    t.RollBack()
    raise
