# -*- coding: utf-8 -*-
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
from revit_duct import RevitDuct
from tag_duct import TagDuct

#.NET Imports
# ==================================================
import clr
clr.AddReference('System')
from System.Collections.Generic import List


# Variables
# ==================================================
app   = __revit__.Application #type: Application
uidoc = __revit__.ActiveUIDocument #type: UIDocument
doc   = revit.doc #type: Document
view  = revit.active_view

ducts = (DB.FilteredElementCollector(doc, view.Id)
         .OfCategory(DB.BuiltInCategory.OST_FabricationDuctwork)
         .WhereElementIsNotElementType()
         .ToElements())

# Main Code
# ==================================================

# Get Views - Selected in a ProjectBroswer
sel_el_ids = uidoc.Selection.GetElementIds()
sel_elem = [doc.GetElement(e_id) for e_id in sel_el_ids]
sel_views = [el for el in sel_elem if issubclass(type(el), View)]

# if none selected - promp selectviews from pyrevit.form.selct_views()
if not sel_views:
    sel_views = forms.select_views()

# Ensure Views Selected
if not sel_views:
    forms.alert("No Views Selected. Please Try Again.", exitscript=True)

# User entered values
# https://revitpythonwrapper.readthedocs.io/en/latest/
from rpw.ui.forms import (FlexForm, Label, TextBox, Separator, Button)
components = [Label("Add refix:"),                      TextBox("prefix"),
              Label("Word you want to replace:"),       TextBox("find"),
              Label("Replacement for that word:"),      TextBox("replace"),
              Label("Add suffix:"),                     TextBox("suffix"),
              Separator(),                              Button("Rename Views")]

form = FlexForm("Rename Views", components)
form.show()

user_inputs = form.values
prefix      = user_inputs["prefix"]
find        = user_inputs["find"]
replace     = user_inputs["replace"]
suffix      = user_inputs["suffix"]


# start transaction to make changes in project
t = Transaction(doc, "py-Rename Views")

t.Start() 

for view in sel_views:

    # create new view name
    old_name = view.Name
    new_name = prefix + old_name.replace(find, replace) + suffix

    # rename views - ensure unique view names
    
    for i in range(20):
        try:
            view.Name = new_name
            print("{} -> {}".format(old_name, new_name))
            break
        except:
            new_name += "*"

t.Commit()

print("Done!")