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
__title__ = "Tag BOD / Length / Size"
__doc__ = """
Tag selected elements with BOD, Length, or Size tags
"""

# Variables
# ==================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
output = script.get_output()
view = revit.active_view
tagger = RevitTagging(doc=doc, view=view)

# Define tags and their positions
tag_configs = {
    'BOD': {
        'tags': ['_umi_bod', '-fabduct_bod_mv_tag'],
        'position': 'center'
    },
    'Length': {
        'tags': ['_umi_length', '-fabduct_length_mv_tag'],
        'position': 'end'
    },
    'Size': {
        'tags': ['_umi_size', '-fabduct_size_mv_tag'],
        'position': 'start'
    }
}

# Code
# ==================================================

# Get selected elements
selected_ids = uidoc.Selection.GetElementIds()
if not selected_ids:
    forms.alert("No elements selected. Please select elements to tag.", exitscript=True)

selected_elements = [doc.GetElement(eid) for eid in selected_ids]

t = Transaction(doc, "Tag Selected Elements - BOD, Length and Size")
t.Start()
try:
    # Loop through each tag config (BOD, Length, and Size)
    for tag_choice, config in tag_configs.items():
        tag_to_use = config['tags']
        location_of_tag = config['position']

        # Find the first available tag from tag_to_use list
        tag_label = None
        for tag_name in tag_to_use:
            try:
                tag_label = tagger.get_label(tag_name)
                break
            except LookupError:
                continue

        if not tag_label:
            continue

        # Get tag family name for checking if already tagged
        tag_fam_name = tag_label.Family.Name if tag_label and tag_label.Family else ""

        # Tag each element with this tag type
        for elem in selected_elements:
            try:
                # Check if already tagged with this tag family
                if tagger.already_tagged(elem, tag_fam_name):
                    continue

                # Place tag with rotation
                tag = tagger.place_tag_at_center_with_rotation(
                    elem,
                    tag_label=tag_label,
                    position=location_of_tag
                )

            except Exception as e:
                pass

    t.Commit()
except Exception as e:
    output.print_md("**Transaction error:** {}".format(e))
    t.RollBack()
    script.exit()
