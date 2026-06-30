# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import script, revit
from Autodesk.Revit.DB import XYZ, Transaction, ElementTransformUtils

# Button info
# ======================================================================
__title__ = 'Move Scope Box'
__doc__ = 'Moves selected scope box to new coordinates'

# Variables
# ======================================================================
doc = revit.doc  # type: ignore[attr-defined]
output = script.get_output()

elements = revit.get_selection().elements

if not elements:
    output.print_md('No elements selected.')
else:
    # Move 10 units in X direction (change these values as needed)
    move_vector = XYZ(10, 0, 0)

    with Transaction(doc, 'Move Scope Box') as txn:
        txn.Start()

        for el in elements:
            try:
                # Move the element using ElementTransformUtils
                ElementTransformUtils.MoveElement(doc, el.Id, move_vector)
                output.print_md('✓ Moved: {}'.format(el.Name))

            except Exception as e:
                output.print_md(
                    '✗ Error moving {}: {}'.format(el.Name, str(e)))

        txn.Commit()
