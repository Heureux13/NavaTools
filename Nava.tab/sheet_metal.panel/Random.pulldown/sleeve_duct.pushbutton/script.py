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
from config.parameters_registry import *

# Button display information
# =================================================
__title__ = "Set Sleeve Type"
__doc__ = """
Sets `_type` to `sleeve` on selected elements
"""

# Parameters to set
parameters_to_set = {
    PYT_SLEEVE_VALUE
}
set_value = "sleeve"

# Code
# ==================================================
uidoc = __revit__.ActiveUIDocument
doc = revit.doc

# Get selected elements
selected_ids = uidoc.Selection.GetElementIds()
if not selected_ids:
    forms.alert("Please select one or more elements", exitscript=True)

t = Transaction(doc, "Set Sleeve Type Parameter")
t.Start()
try:
    for elem_id in selected_ids:
        elem = doc.GetElement(elem_id)
        if elem is None:
            continue

        # Set parameters only if they exist on the element
        for param_name in parameters_to_set:
            try:
                param = elem.LookupParameter(param_name)
                if param is None or param.IsReadOnly:
                    continue

                param.Set(set_value)
            except Exception:
                pass

    t.Commit()
except Exception as e:
    t.RollBack()
    raise
