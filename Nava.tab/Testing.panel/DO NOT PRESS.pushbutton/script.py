# -*- coding: utf-8 -*-
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""


__title__   = "DO NOT PRESS"
__doc__     = """
****************************************************************
Description:

This button is for testin code snippets and ideas before implementing them.
Odds are it will be constantly changing and not useful, its entire purpose
is for the author to have a quick button to test whatever code they are working on.
If you press it could do nothing, throw an error, or change something in your model.
Once working, it will most likely be moved to a more permanent location.

Current goal fucntion of button is: Rename views but not sure how just yet.
****************************************************************
"""

# Imports
# ==================================================
from Autodesk.Revit.DB import *
from pyrevit import revit, forms, DB
from Autodesk.Revit.UI import UIDocument
from Autodesk.Revit.ApplicationServices import Application
from revit_duct import RevitDuct, JointSize
from tag_duct import TagDuct

#.NET Imports
# ==================================================
import clr
clr.AddReference('System')
from System.Collections.Generic import List


# Variables
# ==================================================
app   = __revit__.Application           #type: Application
uidoc = __revit__.ActiveUIDocument      #type: UIDocument
doc   = revit.doc                       #type: Document
view  = revit.active_view

ducts = (DB.FilteredElementCollector(doc, view.Id)
         .OfCategory(DB.BuiltInCategory.OST_FabricationDuctwork)
         .WhereElementIsNotElementType()
         .ToElements())

# Main Code
# ==================================================

ducts   = [RevitDuct(doc, view, el) for el in ducts]
shorts  = [d for d in ducts if d.is_full_joint == JointSize.SHORT]

short_ids   = []
for short in shorts:
    short_ids.append(short.element.Id)

if short_ids:
    id_list = List[ElementId](short_ids)
    uidoc.Selection.SetElementIds(id_list)
    forms.alert("Selected {} short joints".format(len(short_ids)))
else:
    forms.alert("No short joints found")

print("It's called a do not press button for a reason, why did you press it?")