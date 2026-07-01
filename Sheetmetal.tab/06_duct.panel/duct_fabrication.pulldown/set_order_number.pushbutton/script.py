#
# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.DB import Transaction
from pyrevit import revit, script
from ducts.revit_duct import RevitDuct
from revit.revit_element import RevitElement
from config.parameters_registry import (
    PYT_NUMBER_ORDER,
)

# Button info
# ======================================================================
__title__ = 'Set Order Number'
__doc__ = '''
Sets a new order number based on previous numbers.
'''

# Variables
# ======================================================================

output = script.get_output()
doc = revit.doc
uidoc = revit.uidoc
view = revit.view
selected_ducts = RevitDuct.from_selection(uidoc, doc)

ducts = RevitDuct.all(doc)

found = []

for d in ducts:
    number = d._get_param(PYT_NUMBER_ORDER)

    if number and number.strip():
        found.append(number)

ints = []
for x in found:
    try:
        ints.append(int(x))
    except Exception:
        pass

highest = max(ints) if ints else 0
new_num = highest + 1

txn = Transaction(doc, "New Order Number")
txn.Start()

try:
    if not selected_ducts:
        output.print_md("No ducts selected")
    else:
        for d in selected_ducts:
            success = RevitElement(doc, view, d.element).set_param(
                PYT_NUMBER_ORDER, str(new_num))
            if success:
                pass
            else:
                output.print_md('failed to set duct {}'.format(d.id))

    txn.Commit()
except Exception as ex:
    txn.RollBack()

# General debugging prints

# output.print_md('Testing script is running.')
# output.print_md("number of ducts: {}".format(len(ducts)))
# output.print_md("number of order number ducts: {}".format(len(found)))
# output.print_md("Highest number is {}".format(highest))
# output.print_md("new number: {}".format(new_num))
