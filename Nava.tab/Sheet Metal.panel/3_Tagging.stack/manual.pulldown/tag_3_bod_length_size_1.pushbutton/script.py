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
from tagging.revit_tagging import RevitTagging
from revit_element import RevitElement
from revit_duct import RevitDuct
from revit_xyz import RevitXYZ
from pyrevit import DB, forms, revit, script
from Autodesk.Revit.DB import ElementId, Transaction

# Button info
# ==================================================
__title__ = "3 Lenght/BOD/Size 1"
__doc__ = """
Tag selected elements with BOD, Length, and Size tags
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
    'Length': {
        'tags': ['_umi_length_left',],
        'position': 'start'
    },
    'BOD': {
        'tags': ['_umi_bod_center',],
        'position': 'center'
    },
    'Size': {
        'tags': ['_umi_size_right',],
        'position': 'end'
    }
}

# Code
# ==================================================

# Get selected elements
selected_ids = uidoc.Selection.GetElementIds()
if not selected_ids:
    forms.alert(
        "No elements selected. Please select elements to tag.", exitscript=True)

selected_elements = [doc.GetElement(eid) for eid in selected_ids]
# Build a one-pass cache of existing tags in this view to avoid repeated full scans.
existing_tagging.tag_map = tagger.build_existing_tag_family_map(
    selected_elements)

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
            output.print_md(
                "**Skipped {}:** no matching tag family found from {}".format(tag_choice, ", ".join(tag_to_use)))
            continue

        # Get tag family name for checking if already tagged
        tag_fam_name = tag_label.Family.Name if tag_label and tag_label.Family else ""
        tag_fam_name_norm = tag_fam_name.strip().lower()

        # Tag each element with this tag type
        for elem in selected_elements:
            try:
                # Check if already tagged with this tag family
                elem_key = elem.Id.IntegerValue if elem and elem.Id else None
                existing_fams = existing_tagging.tag_map.get(
                    elem_key, set()) if elem_key is not None else set()
                if tag_fam_name_norm and tag_fam_name_norm in existing_fams:
                    continue

                # Place tag with rotation
                tag = tagger.place_tag_at_center_with_rotation(
                    elem,
                    tag_label=tag_label,
                    position=location_of_tag
                )
                if tag and elem_key is not None and tag_fam_name_norm:
                    if elem_key not in existing_tagging.tag_map:
                        existing_tagging.tag_map[elem_key] = set()
                    existing_tagging.tag_map[elem_key].add(tag_fam_name_norm)

            except Exception as e:
                elem_id = elem.Id.IntegerValue if elem and elem.Id else "Unknown"
                output.print_md(
                    "**Skipped {} tag on element {}:** {}".format(tag_choice, elem_id, e))

    t.Commit()
except Exception as e:
    output.print_md("**Transaction error:** {}".format(e))
    t.RollBack()
    script.exit()
