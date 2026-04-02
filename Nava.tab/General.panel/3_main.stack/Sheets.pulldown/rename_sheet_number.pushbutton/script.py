# -*- coding: utf-8 -*-
from rpw.ui.forms import (FlexForm, Label, TextBox, Separator, Button)
from System.Collections.Generic import List
import clr
from Autodesk.Revit.ApplicationServices import Application
from Autodesk.Revit.UI import UIDocument
from pyrevit import revit, forms, DB
from Autodesk.Revit.DB import *
__title__ = "Rename Sheet #s"
__doc__ = """
****************************************************************
Description:

Select sheets to rename. You can then give them a prefix and/or suffix,
and also replace text in the sheet numbers with other text.
****************************************************************
"""


app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = revit.doc

# Get Sheets - Selected in Project Browser
sel_el_ids = uidoc.Selection.GetElementIds()
sel_elem = [doc.GetElement(e_id) for e_id in sel_el_ids]
sel_sheets = [el for el in sel_elem if isinstance(el, ViewSheet)]

# If none selected - prompt select sheets from pyrevit.forms.SelectFromList
if not sel_sheets:
    all_sheets = list(FilteredElementCollector(doc).OfClass(ViewSheet))
    for sheet in all_sheets:
        sheet.display_name = "{} - {}".format(sheet.SheetNumber, sheet.Name)
    sel_sheets = forms.SelectFromList.show(
        sorted(all_sheets, key=lambda s: s.SheetNumber),
        name_attr='display_name',
        multiselect=True,
        title='Select Sheets to Rename'
    )

# Ensure Sheets Selected
if not sel_sheets:
    forms.alert("No Sheets Selected. Please Try Again.", exitscript=True)

components = [Label("Add prefix:"), TextBox("prefix"),
              Label("Text you want to replace:"), TextBox("find"),
              Label("Replacement for that text:"), TextBox("replace"),
              Label("Add suffix:"), TextBox("suffix"),
              Separator(), Button("Rename Sheet Numbers")]

form = FlexForm("Rename Sheet Numbers", components)
form.show()

user_inputs = form.values
if not user_inputs:
    forms.alert("Rename canceled.", exitscript=True)

prefix = user_inputs["prefix"]
find = user_inputs["find"]
replace = user_inputs["replace"]
suffix = user_inputs["suffix"]

# start transaction to make changes in project
t = Transaction(doc, "Rename Sheet Numbers")
t.Start()

for sheet in sel_sheets:
    old_number = sheet.SheetNumber
    new_number = prefix + old_number.replace(find, replace) + suffix

    # rename sheets - ensure unique sheet numbers
    for i in range(20):
        try:
            sheet.SheetNumber = new_number
            print("{} -> {}".format(old_number, new_number))
            break
        except BaseException:
            new_number += "*"

t.Commit()
print("Done!")
