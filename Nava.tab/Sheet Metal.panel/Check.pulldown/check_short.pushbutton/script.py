# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from revit_element import RevitElement
from revit_duct import RevitDuct
from revit_output import print_disclaimer
from tag_slot_config import DEFAULT_SKIP_PARAMETERS
from pyrevit import revit, script
from Autodesk.Revit.DB import *

# Button info
# ===================================================
__title__ = "Short"
__doc__ = """
Selects straight duct that is shorter than 12".
"""

# Variables
# ==================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
view = revit.active_view
output = script.get_output()
MAX_LENGTH_IN = 12.01
SKIP_VALUES = {
    str(v).strip().lower()
    for values in DEFAULT_SKIP_PARAMETERS.values()
    for v in (values or [])
    if v is not None
}


def _get_param_value_from_element_or_type(element, param_name):
    """Return string parameter value from instance or type, or None."""
    if element is None or not param_name:
        return None

    try:
        p = element.LookupParameter(param_name)
    except Exception:
        p = None
    if p:
        try:
            val = p.AsString() or p.AsValueString()
            if val is not None:
                return str(val).strip()
        except Exception:
            pass

    try:
        type_el = doc.GetElement(element.GetTypeId())
    except Exception:
        type_el = None
    if type_el:
        try:
            p = type_el.LookupParameter(param_name)
        except Exception:
            p = None
        if p:
            try:
                val = p.AsString() or p.AsValueString()
                if val is not None:
                    return str(val).strip()
            except Exception:
                pass

    return None


def _should_skip_by_item_number(duct):
    """Skip when Item Number matches configured skip values."""
    item_number = _get_param_value_from_element_or_type(
        duct.element, "Item Number")
    if not item_number:
        return False
    return item_number.strip().lower() in SKIP_VALUES

# Main Code
# ==================================================


# Get all ducts
ducts = RevitDuct.all(doc, view)

# Filter down to straight ducts shorter than threshold
fil_ducts = [
    d for d in ducts
    if (d.family or "").strip().lower() == "straight"
    and d.length < MAX_LENGTH_IN
    and not _should_skip_by_item_number(d)
]

# Start of select / print loop
if fil_ducts:

    # Select filtered dcuts
    RevitElement.select_many(uidoc, fil_ducts)
    output.print_md('# Selected {} straight ducts shorter than {}"'.format(len(fil_ducts), int(MAX_LENGTH_IN)))
    output.print_md("---")

    # Individutal duct and selected properties
    for i, fil in enumerate(fil_ducts, start=1):
        output.print_md(
            '### No: {:03} | ID: {} | Length: {:06.2f}" | Size: {} | Connectors: 1 = {}, 2 = {}'.format(
                i,
                output.linkify(fil.element.Id),
                fil.length,
                fil.size,
                fil.connector_0_type,
                fil.connector_1_type,
            )
        )

    # Total count
    element_ids = [d.element.Id for d in fil_ducts]
    output.print_md(
        "# Total elements {}, {}".format(
            len(element_ids), output.linkify(element_ids))
    )

    # Final print statements
    print_disclaimer(output)
else:
    output.print_md('## No straight ducts shorter than {}" selected'.format(int(MAX_LENGTH_IN)))
