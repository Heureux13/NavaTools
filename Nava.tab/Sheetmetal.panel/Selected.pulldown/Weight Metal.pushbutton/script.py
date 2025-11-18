# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from System.Collections.Generic import List
from revit_duct import RevitDuct
from Autodesk.Revit.ApplicationServices import Application
from Autodesk.Revit.UI import UIDocument
from pyrevit import revit, forms, script
from Autodesk.Revit.DB import *

# Button info
# =================================================
__title__ = "Weight Metal"
__doc__ = """
******************************************************************
Description:
Returns metal weight for duct(s) selected.
******************************************************************
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
ducts = RevitDuct.from_selection(uidoc, doc, view)

if not ducts:
    forms.alert("Please select one or more ducts first.")
else:
    # keep both the ElementId and the weight
    weights = [(d.element.Id, d.id, d.weight_metal)
               for d in ducts if d.weight_metal is not None]

    # Section title
    output.print_md("### Metal Weights")

    # Individual links with weights
    for eid, id_int, w in weights:
        output.print_md("- {}: {:.2f} lbs".format(output.linkify(eid), w))

    # Select All link
    all_ids = List[ElementId]()
    for eid, _, _ in weights:
        all_ids.Add(eid)
    output.print_md("**{}**".format(output.linkify(all_ids)))

    # Footer total
    total = sum(w for _, _, w in weights)
    output.print_md("**➡️ Total Metal weight: {:.2f} lbs**".format(total))
