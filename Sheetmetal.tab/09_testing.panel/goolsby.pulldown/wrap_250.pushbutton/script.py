# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import revit, forms, script
from Autodesk.Revit.DB import StorageType, Transaction
from config.parameters_registry import PYT_NOTE_0

# Button info
# ======================================================================
__title__ = 'Wrap 2.50"'
__doc__ = '''
Set _UMI_PYT_Note0 to Liner 1.00" on selected elements.
'''

# Variables
# ======================================================================

output = script.get_output()
doc = revit.doc
uidoc = __revit__.ActiveUIDocument

NOTE_VALUE = 'Wrap 2.50"'


def _get_param_case_insensitive(element, param_name):
    target = (param_name or '').strip().lower()
    if not target or element is None:
        return None

    try:
        direct = element.LookupParameter(param_name)
        if direct:
            return direct
    except Exception:
        pass

    for param in element.Parameters:
        try:
            definition = param.Definition
            name = definition.Name if definition else None
            if name and name.strip().lower() == target:
                return param
        except Exception:
            pass
    return None


selected_ids = uidoc.Selection.GetElementIds()
if not selected_ids:
    forms.alert('Please select one or more elements.', exitscript=True)

updated = 0
unchanged = 0
missing = 0
readonly = 0
wrong_type = 0

t = Transaction(doc, 'Set _UMI_PYT_Note0')
t.Start()
try:
    for elem_id in selected_ids:
        element = doc.GetElement(elem_id)
        if element is None:
            continue

        param = _get_param_case_insensitive(element, PYT_NOTE_0)
        if param is None:
            missing += 1
            continue

        if param.IsReadOnly:
            readonly += 1
            continue

        if param.StorageType != StorageType.String:
            wrong_type += 1
            continue

        current = (param.AsString() or '').strip()
        if current == NOTE_VALUE:
            unchanged += 1
            continue

        param.Set(NOTE_VALUE)
        updated += 1

    t.Commit()
except Exception:
    t.RollBack()
    raise

# output.print_md('## _UMI_PYT_Note0 update complete')
# output.print_md('- Selected: {}'.format(len(selected_ids)))
# output.print_md('- Updated: {}'.format(updated))
# output.print_md('- Unchanged: {}'.format(unchanged))
# output.print_md('- Missing parameter: {}'.format(missing))
# output.print_md('- Read-only: {}'.format(readonly))
# output.print_md('- Non-text parameter: {}'.format(wrong_type))
