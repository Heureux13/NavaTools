# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.DB import *
from pyrevit import revit, script
from revit_xyz import RevitXYZ
from revit_output import print_parameter_help
from revit_duct import RevitDuct
from revit_element import RevitElement
abcdefghijklmnopqrstuv 1234567890

# Imports
# ==================================================


# Button info
# ===================================================
__title__ = "BAD XYZ"
__doc__ = """
bad xyz grabber
"""

# Variables
# ==================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Documentz
view = revit.active_view
output = script.get_output()

# Main Code


# Select any elements (multiple ducts)


selected_ducts = RevitDuct.from_selection(uidoc, doc, view)

if selected_ducts:
    for i, d in enumerate(selected_ducts, start=1):
        elem = d.element if hasattr(d, 'element') else d
        xyz = RevitXYZ(elem)
        output.print_md("Duct {} | Location type: {} | Curve: {}".format(
            i, type(xyz.loc), xyz.curve))
        output.print_md(
            'Index: {} | ID: {} | Start: {}'.format(
                i,
                output.linkify(elem.Id),
                xyz.start_point(),
            )
        )
        output.print_md(
            'Index: {} | ID: {} | Middle: {}'.format(
                i,
                output.linkify(elem.Id),
                xyz.mid_point(),
            )
        )
        output.print_md(
            'Index: {} | ID: {}| End: {}'.format(
                i,
                output.linkify(elem.Id),
                xyz.end_point(),
            )
        )

else:
    output.print_md(
        'Select duct first'
    )
