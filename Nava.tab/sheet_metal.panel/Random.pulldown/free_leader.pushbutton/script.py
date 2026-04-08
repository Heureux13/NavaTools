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
from Autodesk.Revit.DB import Transaction, IndependentTag, BuiltInCategory

# Button display information
# =================================================
__title__ = "Free Leader"
__doc__ = """
Take what ever annotation you have selected and it will add leader line and make it free end.
"""

# Code
# ==================================================
uidoc = __revit__.ActiveUIDocument
doc = revit.doc

# Get selected elements
selected_ids = uidoc.Selection.GetElementIds()
if not selected_ids:
    import sys
    sys.exit()

processed_count = 0
skipped_count = 0

t = Transaction(doc, "Add Free Leader to Annotations")
t.Start()
try:
    for elem_id in selected_ids:
        elem = doc.GetElement(elem_id)
        if elem is None:
            skipped_count += 1
            continue

        success = False

        # Try using parameters first (works for MEP Fabrication tags)
        leader_line_param = elem.LookupParameter("Leader Line")
        leader_type_param = elem.LookupParameter("Leader Type")

        if leader_line_param and leader_type_param:
            try:
                # Step 1: Enable leader line first
                if not leader_line_param.IsReadOnly:
                    # Check current value and enable if not already
                    current_value = leader_line_param.AsValueString()
                    if current_value and current_value.strip().lower() == "no":
                        leader_line_param.Set(1)

                # Regenerate to update the document state
                doc.Regenerate()

                # Step 2: Now set to "Free End" after leader is enabled
                if not leader_type_param.IsReadOnly:
                    # Check if we need to change the value
                    current_type = leader_type_param.AsValueString()
                    if not current_type or current_type.strip().lower() != "free end":
                        # Try setting as integer first (1 might be Free End, 0 might be Attached End)
                        try:
                            leader_type_param.Set(1)
                        except BaseException:
                            # If integer doesn't work, try string
                            try:
                                leader_type_param.SetValueString("Free End")
                            except BaseException:
                                # Last resort: try 0
                                pass

                processed_count += 1
                success = True
            except Exception:
                pass

        # Fall back to IndependentTag API for standard Revit tags
        if not success and isinstance(elem, IndependentTag):
            try:
                # Enable leader if not already enabled
                if not elem.HasLeader:
                    elem.HasLeader = True

                # Set leader end condition to free
                elem.LeaderEndCondition = 0  # 0 = Free, 1 = Attached

                processed_count += 1
                success = True
            except Exception:
                pass

        if not success:
            skipped_count += 1

    t.Commit()

except Exception as e:
    t.RollBack()
    forms.alert("Error: {}".format(str(e)))
    raise
