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
from constants.print_outputs import print_disclaimer
from tagging.revit_tagging import RevitTagging
from tagging.tag_config import DEFAULT_TAG_SLOT_CANDIDATES, SLOT_LENGTH
from revit.revit_element import RevitElement
from ducts.revit_duct import RevitDuct
from ducts.revit_xyz import RevitXYZ
from pyrevit import DB, forms, revit, script
from Autodesk.Revit.DB import ElementId, Transaction

# Button info
# ==================================================
__title__ = "1 length"
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

# Helper function for version-compatible ElementId access


def get_element_id_value(element_id):
    """Extract integer value from ElementId (compatible with Revit 2025 and 2026+)."""
    if element_id is None:
        return None
    try:
        # Revit 2025 and earlier
        return element_id.IntegerValue
    except AttributeError:
        # Revit 2026+ uses .Value
        try:
            return element_id.Value
        except (AttributeError, TypeError):
            # Fallback: try direct int conversion
            return int(element_id)


tag_to_use = list(DEFAULT_TAG_SLOT_CANDIDATES.get(SLOT_LENGTH, []))

location_of_tag = 'center'

# Code
# ==================================================

# Get selected elements
selected_ids = uidoc.Selection.GetElementIds()
if not selected_ids:
    forms.alert(
        "No elements selected. Please select elements to tag.", exitscript=True)

selected_elements = [doc.GetElement(eid) for eid in selected_ids]
existing_tag_map = tagger.build_existing_tag_family_map(
    selected_elements)

# Find the first available tag from tag_to_use list
tag_label = None
for tag_name in tag_to_use:
    try:
        if isinstance(tag_name, tuple):
            tag_label = tagger.get_label_exact(tag_name[0], tag_name[1])
        else:
            tag_label = tagger.get_label(tag_name)
        break
    except LookupError:
        continue

if not tag_label:
    missing = []
    for name in tag_to_use:
        if isinstance(name, tuple):
            missing.append("{} :: {}".format(name[0], name[1]))
        else:
            missing.append(str(name))
    forms.alert(
        "None of the specified tags were found:\n{}".format(
            "\n".join(missing)),
        exitscript=True
    )

# Tag selected elements
placed = []
failed = []
already_tagged = []

# Get tag family name for checking if already tagged
tag_fam_name = tag_label.Family.Name if tag_label and tag_label.Family else ""
tag_fam_name_norm = tag_fam_name.strip().lower()

t = Transaction(doc, "Tag Selected Elements")
t.Start()
try:
    for elem in selected_elements:
        try:
            # Check if already tagged with this tag family
            elem_key = get_element_id_value(elem.Id)
            existing_fams = existing_tag_map.get(
                elem_key, set()) if elem_key is not None else set()
            if tag_fam_name_norm and tag_fam_name_norm in existing_fams:
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
                if elem_key is not None and tag_fam_name_norm:
                    if elem_key not in existing_tag_map:
                        existing_tag_map[elem_key] = set()
                    existing_tag_map[elem_key].add(tag_fam_name_norm)
            else:
                failed.append((elem, "Tag placement returned None"))

        except Exception as e:
            failed.append((elem, "Error: {}".format(str(e))))

    t.Commit()
except Exception as e:
    output.print_md("**Transaction error:** {}".format(e))
    t.RollBack()
    script.exit()
