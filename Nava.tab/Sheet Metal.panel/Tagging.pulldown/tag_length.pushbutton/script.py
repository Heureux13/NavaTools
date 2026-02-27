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
from revit_output import print_disclaimer
from revit_tagging import RevitTagging
from revit_element import RevitElement
from revit_duct import RevitDuct
from revit_xyz import RevitXYZ
from pyrevit import DB, forms, revit, script
from Autodesk.Revit.DB import ElementId, Transaction

# Button info
# ==================================================
__title__ = "Tag length"
__doc__ = """
Tags with length tag at end of duct
"""

# Variables
# ==================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
output = script.get_output()
view = revit.active_view
tagger = RevitTagging(doc=doc, view=view)

tag_to_use = [
    '_umi_length',
    '-fabduct_length_mv_tag',
]

location_of_tag = 'end'

# Code
# ==================================================

# Get selected elements
selected_ids = uidoc.Selection.GetElementIds()
if not selected_ids:
    forms.alert("No elements selected. Please select elements to tag.", exitscript=True)

selected_elements = [doc.GetElement(eid) for eid in selected_ids]

# Find the first available tag from tag_to_use list
tag_label = None
for tag_name in tag_to_use:
    try:
        tag_label = tagger.get_label(tag_name)
        break
    except LookupError:
        continue

if not tag_label:
    forms.alert(
        "None of the specified tags were found:\n{}".format("\n".join(tag_to_use)),
        exitscript=True
    )

# Tag selected elements
placed = []
failed = []
already_tagged = []

# Get tag family name for checking if already tagged
tag_fam_name = tag_label.Family.Name if tag_label and tag_label.Family else ""

t = Transaction(doc, "Tag Selected Elements")
t.Start()
try:
    for elem in selected_elements:
        try:
            # Check if already tagged with this tag family
            if tagger.already_tagged(elem, tag_fam_name):
                already_tagged.append(elem)
                continue

            # Place tag with rotation
            tag = tagger.place_tag_at_center_with_rotation(
                elem,
                tag_label=tag_label,
                position=location_of_tag
            )

            if tag:
                placed.append(elem)
            else:
                failed.append((elem, "Tag placement returned None"))

        except Exception as e:
            failed.append((elem, "Error: {}".format(str(e))))

    t.Commit()
except Exception as e:
    output.print_md("**Transaction error:** {}".format(e))
    t.RollBack()
    script.exit()
