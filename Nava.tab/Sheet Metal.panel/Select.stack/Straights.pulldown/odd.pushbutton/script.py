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
from revit_output import print_parameter_help
from revit_duct import RevitDuct
from pyrevit import revit, script
from Autodesk.Revit.DB import *

# Button info
# ===================================================
__title__ = "Joints Odd"
__doc__ = """
Selects all spiral joints that are odd size
"""

# Variables
# ==================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
view = revit.active_view
output = script.get_output()

# Main Code
# ==================================================

# Get all duct
all_ducts = RevitDuct.all(doc, view)

# Families allowed
allowed = {
    ("straight", "tdf"),
    ("straight", "s&d"),
    ("spiral duct", "raw"),
    ("boot tap - wdamper", "bead"),
    ("boot saddle tap", "flange out - 1in"),
    ("45 tap", "bead & slip"),
    ("coupling - fitting", "raw"),
    ('gored elbow', 'bead'),
    ('reducer', 'bead'),
}

# Nomalize and filter duct
normalized = [(d, (d.family or "").lower().strip(),
               (d.connector_0_type or "").lower().strip()) for d in all_ducts]

fil_ducts = [d for d, fam, conn in normalized if (fam, conn) in allowed]

short_ducts = [
    d for d in fil_ducts
    if (
        d.size and
        ('Ø' in d.size or 'ø' in d.size) and
        any(
            num.isdigit() and int(num) % 2 == 1
            for num in [s.split('"')[0] for s in d.size.split('-') if ('Ø' in s or 'ø' in s)]
        )
    )
]

sd = short_ducts

# Start of select / print
if sd:

    # Select filtered duct
    RevitElement.select_many(uidoc, sd)
    output.print_md(
        "# Selected {} odd size joints".format(len(sd))
    )
    output.print_md("---")

    # Individual duct and properties
    for i, fil in enumerate(sd, start=1):
        length_in = fil.length or 0.0
        output.print_md(
            '### No: {:03} | ID: {} | Size: {} | Family: {}'.format(
                i,
                output.linkify(fil.element.Id),
                fil.size,
                fil.family
            )
        )

    element_ids = [d.element.Id for d in sd]
    output.print_md(
        "# Total elements {}, {}".format(
            len(sd), output.linkify(element_ids)
        )
    )

    # Final print statements
    print_parameter_help(output)
else:
    output.print_md("## No odd size joints found")
