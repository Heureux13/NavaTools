# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import revit, script
from Autodesk.Revit.DB import StorageType

# Button info
# ======================================================================
__title__ = 'Reverse Offset'
__doc__ = '''
Reverses the offset direction by:
- Swapping DN ↔ UP
- Flipping arrows 180 degrees: → ← ↓ ↑
- Keeping numbers intact
'''

# Variables
# ======================================================================

output = script.get_output()


def reverse_offset_value(value):
    """Reverse the offset value by swapping DN/UP and flipping arrows."""
    if not value:
        return value

    result = value

    # Swap DN and UP using a placeholder
    result = result.replace('DN', '\x00TEMP_DN\x00')
    result = result.replace('UP', 'DN')
    result = result.replace('\x00TEMP_DN\x00', 'UP')

    # Flip arrows 180 degrees
    result = result.replace('→', '\x00ARROW_RIGHT\x00')
    result = result.replace('←', '→')
    result = result.replace('\x00ARROW_RIGHT\x00', '←')

    result = result.replace('↓', '\x00ARROW_UP\x00')
    result = result.replace('↑', '↓')
    result = result.replace('\x00ARROW_UP\x00', '↑')

    return result


# Main Code
# ======================================================================

# Get selected elements
selection = revit.get_selection()
doc = revit.doc

if not selection:
    output.print_md("# Please select ductwork and try again")
else:
    output.print_md("# Reverse Offset")

    updated_count = 0

    with revit.Transaction("Reverse Offset"):
        for element in selection:
            try:
                # Get the _duct_tag_offset parameter
                tag_p = element.LookupParameter('_duct_tag_offset')

                if tag_p and not tag_p.IsReadOnly:
                    if tag_p.StorageType == StorageType.String:
                        current_value = tag_p.AsString()

                        if current_value and current_value.strip():
                            # Reverse the value
                            reversed_value = reverse_offset_value(current_value)
                            tag_p.Set(reversed_value)
                            updated_count += 1

                            output.print_md(
                                "### ID: {} | {} → {}".format(
                                    output.linkify(element.Id),
                                    current_value,
                                    reversed_value
                                ))
            except Exception as e:
                output.print_md("ERROR: processing element {} : {}".format(
                    element.Id.Value,
                    str(e)
                ))

    output.print_md("---")
    output.print_md("# Summary: {} elements updated".format(updated_count))
