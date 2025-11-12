# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

__title__   = "DO NOT PRESS"
__doc__     = """
******************************************************************
Description:

Current goal fucntion of button is: select only spiral duct.

++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
This button is for testin code snippets and ideas before implementing them.
Odds are it will be constantly changing and not useful, its entire purpose
is for the author to have a quick button to test whatever code they are working on.
If you press it could do nothing, throw an error, or change something in your model.
Once working, it will most likely be moved to a more permanent location.
******************************************************************
"""

# Imports
# ==================================================
from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script, DB
from Autodesk.Revit.UI import UIDocument
from Autodesk.Revit.ApplicationServices import Application
from revit_duct import RevitDuct, JointSize, CONNECTOR_THRESHOLDS
from tag_duct import TagDuct
from revit_element import RevitElement
from revit_parameter import RevitParameter
from revit_tagging import RevitTagging, RevitXYZ

#.NET Imports
# ==================================================
import clr
from System.Collections.Generic import List


# Variables
# ==================================================
app   = __revit__.Application           #type: Application
uidoc = __revit__.ActiveUIDocument      #type: UIDocument
doc   = revit.doc                       #type: Document
view  = revit.active_view
output = script.get_output()

# Main Code
# ==================================================
ducts = RevitDuct.all(doc, view)
fil_ducts  = [d for d in ducts if d.joint_size == JointSize.SHORT]


for i, el in enumerate(fil_ducts):
    loc = getattr(el, "Location", None)
    has_curve = getattr(loc, "Curve", None) is not None
    is_loc_point = loc is not None and loc.GetType().Name == "LocationPoint"
    bb = el.get_BoundingBox(None)
    print(i, "Id:", el.Id.IntegerValue, "Category:", getattr(el, "Category", None).Name if getattr(el, "Category", None) else None,
          "HasLocation:", loc is not None,
          "HasCurve:", has_curve,
          "IsLocationPoint:", is_loc_point,
          "HasBBox:", bb is not None)
    
# # collect X for each filtered duct (skip ducts without curve)
# xs = []
# for el in fil_ducts:
#     loc_obj = getattr(el, "Location", None)
#     curve = getattr(loc_obj, "Curve", None)
#     if curve:
#         xs.append(curve.GetEndPoint(0).X)
#     else:
#         xs.append(None)

# # select all filtered ducts (pass the list, not a single element)
# RevitElement.select_many(uidoc, fil_ducts)

# # output results
# print(xs)