# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
from xml.dom.minidom import Document

# ======================================================================

from pyrevit import script, forms
from Autodesk.Revit.DB import *
from Autodesk.Revit.Exceptions import ArgumentException
from rpw.ui.forms import FlexForm, Label, TextBox, Separator, Button


# Button info
# ======================================================================
__title__ = 'Ginger Rename'
__doc__ = '''
Follow the instructions
'''

# Variables
# ======================================================================
doc = __revit__.ActiveUIDocument.Document #type:Document
app = __revit__.Application
output = script.get_output() # pyRevit Output Menu

def get_rename_parameters():
    components = [
        Label("Prefix:"),   TextBox("prefix",   Text=""),
        Label("Find:"),     TextBox("find",     Text=""),
        Label("Replace:"),  TextBox("replace",  Text=""),
        Label("Suffix:"),   TextBox("suffix",   Text=""),
        Separator(),
        Button("OK"),
    ]
    form = FlexForm("Name Swapper", components)
    form.show()
    values = form.values

    if not values:
        forms.alert("No Values Entered", exitscript=True)

    PREFIX  = values["prefix"]
    FIND    = values["find"]
    REPLACE = values["replace"]
    SUFFIX  = values["suffix"]

    return PREFIX, FIND, REPLACE, SUFFIX

# 1. Select Views
selected_views = forms.select_views()

if not selected_views:
    forms.alert("No Views Selected", exitscript=True)

# 2. Define renaming rules
PREFIX, FIND, REPLACE, SUFFIX = get_rename_parameters()

# 3. Change view names
name_changes = []
dont_allow   = r'\\:{\}[]|;<>?`~'

t = Transaction(doc, "Name Swapper")
t.Start()
try:
    for view in selected_views:
        current_name = view.Name
        renamed_name = current_name.replace(FIND, REPLACE) if FIND else current_name
        new_name = PREFIX + renamed_name + SUFFIX
        new_name = ''.join([c for c in new_name if c not in dont_allow])

        renamed = False
        for i in range(10):
            try:
                view.Name = new_name
                renamed = True
                break
            except ArgumentException:
                new_name += '*'

        if not renamed:
            raise ArgumentException("Unable to rename view '{}' after 10 attempts.".format(current_name))

        link = output.linkify(view.Id, "Link")
        name_changes.append([link, current_name, new_name])

    t.Commit()
finally:
    if t.GetStatus() == TransactionStatus.Started:
        t.RollBack()

# Output as pyRevit table
if name_changes:
    output.print_table(
        name_changes,
        title="View Name Changes",
        columns=["View", "Original Name", "New Name"]
    )
