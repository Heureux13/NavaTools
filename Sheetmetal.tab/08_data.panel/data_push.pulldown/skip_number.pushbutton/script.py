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
__title__ = "Skip Number"
__doc__ = """
Sets parameters to 'skip' for selected elements
"""

# Parameters to set to 'skip'
parameters_to_skip = {
    PYT_SKIP_NUMBER,
}

# Code
# ==================================================
uidoc = __revit__.ActiveUIDocument
doc = revit.doc

# Get selected elements
selected_ids = uidoc.Selection.GetElementIds()
if not selected_ids:
    forms.alert("Please select one or more elements", exitscript=True)

t = Transaction(doc, "Toggle Parameters Skip")
t.Start()
try:
    for elem_id in selected_ids:
        elem = doc.GetElement(elem_id)
        if elem is None:
            continue

        # Toggle parameters: set to 'skip' if not already, clear if already 'skip'
        for param_name in parameters_to_skip:
            try:
                param = elem.LookupParameter(param_name)
                if param is None or param.IsReadOnly:
                    continue

                current = (param.AsString() or '').strip().lower()
                if current == 'skip':
                    param.Set('')
                else:
                    param.Set('skip')
            except Exception:
                pass

    t.Commit()
except Exception as e:
    t.RollBack()
    raise
