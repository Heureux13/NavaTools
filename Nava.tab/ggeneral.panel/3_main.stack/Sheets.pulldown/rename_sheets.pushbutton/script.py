# -*- coding: utf-8 -*-
__title__   = "Rename Sheets"
__doc__     = """
****************************************************************
Description:

Select sheets to rename. You can then give them a prefix and/or suffix,
and also replace a word in the sheet names with another word.
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

# Get Sheets - Selected in Project Browser
sel_el_ids = uidoc.Selection.GetElementIds()
sel_elem = [doc.GetElement(e_id) for e_id in sel_el_ids]
sel_sheets = [el for el in sel_elem if isinstance(el, ViewSheet)]

# If none selected - prompt select sheets from pyrevit.forms.SelectFromList
if not sel_sheets:
    all_sheets = list(FilteredElementCollector(doc).OfClass(ViewSheet))
    sel_sheets = forms.SelectFromList.show(
        sorted(all_sheets, key=lambda s: s.Name),
        name_attr='Name',
        multiselect=True,
        title='Select Sheets to Rename'
    )

# Ensure Sheets Selected
if not sel_sheets:
    forms.alert("No Sheets Selected. Please Try Again.", exitscript=True)

from rpw.ui.forms import (FlexForm, Label, TextBox, Separator, Button)
components = [Label("Add prefix:"),                      TextBox("prefix"),
              Label("Word you want to replace:"),       TextBox("find"),
              Label("Replacement for that word:"),      TextBox("replace"),
              Label("Add suffix:"),                     TextBox("suffix"),
              Separator(),                              Button("Rename Sheets")]

form = FlexForm("Rename Sheets", components)
form.show()

user_inputs = form.values
prefix      = user_inputs["prefix"]
find        = user_inputs["find"]
replace     = user_inputs["replace"]
suffix      = user_inputs["suffix"]

# start transaction to make changes in project
t = Transaction(doc, "Rename Sheets")
t.Start()

for sheet in sel_sheets:
    old_name = sheet.Name
    new_name = prefix + old_name.replace(find, replace) + suffix

    # rename sheets - ensure unique sheet names
    for i in range(20):
        try:
            sheet.Name = new_name
            print("{} -> {}".format(old_name, new_name))
            break
        except:
            new_name += "*"

t.Commit()
print("Done!")