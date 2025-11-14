# -*- coding: utf-8 -*-
__title__   = "Rename Families"
__doc__     = """
****************************************************************
Description:

Select families to rename. You can then give them a prefix and/or suffix.
****************************************************************
"""

from Autodesk.Revit.DB import *
from pyrevit import revit, forms, DB
from Autodesk.Revit.UI import UIDocument
from Autodesk.Revit.ApplicationServices import Application
import clr
from System.Collections.Generic import List

app   = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc   = revit.doc

# Get Families - Selected in Project Browser
sel_el_ids = uidoc.Selection.GetElementIds()
sel_elem = [doc.GetElement(e_id) for e_id in sel_el_ids]
sel_fams = [el for el in sel_elem if isinstance(el, Family)]

# If none selected - prompt select families from pyrevit.forms.SelectFromList
if not sel_fams:
    all_fams = list(FilteredElementCollector(doc).OfClass(Family))
    sel_fams = forms.SelectFromList.show(
        sorted(all_fams, key=lambda f: f.Name),
        name_attr='Name',
        multiselect=True,
        title='Select Families to Rename'
    )

# Ensure Families Selected
if not sel_fams:
    forms.alert("No Families Selected. Please Try Again.", exitscript=True)

from rpw.ui.forms import (FlexForm, Label, TextBox, Separator, Button)
components = [Label("Add prefix:"),                      TextBox("prefix"),
              Label("Word you want to replace:"),       TextBox("find"),
              Label("Replacement for that word:"),      TextBox("replace"),
              Label("Add suffix:"),                     TextBox("suffix"),
              Separator(),                              Button("Rename Families")]

form = FlexForm("Rename Families", components)
form.show()

user_inputs = form.values
prefix      = user_inputs["prefix"]
find        = user_inputs["find"]
replace     = user_inputs["replace"]
suffix      = user_inputs["suffix"]

# start transaction to make changes in project
t = Transaction(doc, "Rename Families")
t.Start()

for f in sel_fams:
    old_name = f.Name
    new_name = prefix + old_name.replace(find, replace) + suffix

    # rename families - ensure unique family names
    for i in range(20):
        try:
            f.Name = new_name
            print("{} -> {}".format(old_name, new_name))
            break
        except:
            new_name += "*"

t.Commit()
print("Done!")