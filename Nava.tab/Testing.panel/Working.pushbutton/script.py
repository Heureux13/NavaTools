# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from System.Collections.Generic import List
from Autodesk.Revit.ApplicationServices import Application
from Autodesk.Revit.DB import Transaction, Reference, ElementId
from Autodesk.Revit.UI import UIDocument
from pyrevit import revit, forms, DB, script
from revit_element import RevitElement
from revit_duct import RevitDuct, JointSize, DuctAngleAllowance
from revit_xyz import RevitXYZ
from revit_tagging import RevitTagging
import clr

# Button info
# ==================================================
__title__ = "C XYZ"
__doc__ = """
correct xyz grabber
"""

# Variables
# ==================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
output = script.get_output()
view = revit.active_view
tagger = RevitTagging(doc=doc, view=view)


# Main Code
try:
    ref = uidoc.Selection.PickObject(ObjectType.Element, "Select an element to print XYZ")
    elem = doc.GetElement(ref.ElementId)

    # Try location point
    loc = getattr(elem, "Location", None)
    if loc and hasattr(loc, "Point") and loc.Point:
        pt = loc.Point
        output.print_md("**XYZ (Location.Point):** {}".format(pt))
    # Try curve endpoints
    elif loc and hasattr(loc, "Curve") and loc.Curve:
        start = loc.Curve.GetEndPoint(0)
        end = loc.Curve.GetEndPoint(1)
        output.print_md("**Start XYZ (Curve):** {}".format(start))
        output.print_md("**End XYZ (Curve):** {}".format(end))
    # Try bounding box
    elif elem.get_BoundingBox(None):
        bb = elem.get_BoundingBox(None)
        output.print_md("**XYZ (BoundingBox.Min):** {}".format(bb.Min))
    else:
        output.print_md("**No XYZ found for element.**")
except Exception as e:
    output.print_md("**Error:** {}".format(e))
