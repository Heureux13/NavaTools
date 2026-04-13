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
from tagging.tag_config import DEFAULT_TAG_SLOT_CANDIDATES, SLOT_SIZE_RIGHT, SLOT_BOD_LEFT
from revit.revit_element import RevitElement
from ducts.revit_duct import RevitDuct
from ducts.revit_xyz import RevitXYZ
from pyrevit import DB, forms, revit, script
from Autodesk.Revit.DB import ElementId, Transaction

# Button info
# ==================================================
__title__ = "Tag Sleeves 0"
__doc__ = """
Tag all sleeve ducts in active view with BOD/Size tags
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


# Define tags and their positions
tag_configs = {
    'Length': {
        'tags': list(DEFAULT_TAG_SLOT_CANDIDATES.get(SLOT_SIZE_RIGHT, [])),
        'position': 'start'
    },
    'Size': {
        'tags': list(DEFAULT_TAG_SLOT_CANDIDATES.get(SLOT_BOD_LEFT, [])),
        'position': 'end'
    }
}

# Code
# ==================================================

# Get all fabrication ducts in active view, then keep only sleeves.
all_view_ducts = list(
    DB.FilteredElementCollector(doc, view.Id)
    .OfCategory(DB.BuiltInCategory.OST_FabricationDuctwork)
    .WhereElementIsNotElementType()
    .ToElements()
)


def is_sleeve_type(element):
    p = element.LookupParameter("_type") if element else None
    if not p:
        return False
    raw = p.AsString() or p.AsValueString() or ""
    value = raw.strip().lower()
    return value in ("sleeve", "sleeves")


selected_elements = [e for e in all_view_ducts if is_sleeve_type(e)]
if not selected_elements:
    forms.alert(
        "No fabrication ducts with _type = sleeve/sleeves were found in the active view.", exitscript=True)

existing_tag_map = tagger.build_existing_tag_family_map(selected_elements)

# Track elements that already had Size/BOD-like annotations before this command runs.
pretagged_size_bod_ids = set(
    elem_id for elem_id, fams in existing_tag_map.items()
    if any(("size" in fam) or ("bod" in fam) for fam in fams)
)

t = Transaction(doc, "Tag Selected Elements - Length and Size")
t.Start()
try:
    # Loop through each tag config (Length and Size)
    for tag_choice, config in tag_configs.items():
        tag_to_use = config['tags']
        location_of_tag = config['position']

        # Find the first available tag from tag_to_use list
        tag_label = None
        for tag_name in tag_to_use:
            try:
                if isinstance(tag_name, tuple):
                    tag_label = tagger.get_label_exact(
                        tag_name[0], tag_name[1])
                else:
                    tag_label = tagger.get_label(tag_name)
                break
            except LookupError:
                continue

        if not tag_label:
            continue

        # Get tag family name for checking if already tagged
        tag_fam_name = tag_label.Family.Name if tag_label and tag_label.Family else ""
        tag_fam_name_norm = tag_fam_name.strip().lower()

        # Tag each element with this tag type
        for elem in selected_elements:
            try:
                elem_key = get_element_id_value(elem.Id)
                existing_fams = existing_tag_map.get(
                    elem_key, set()) if elem_key is not None else set()

                # Skip elements that already have a Size or BOD annotation family.
                if elem_key in pretagged_size_bod_ids:
                    continue

                # Check if already tagged with this exact tag family.
                if tag_fam_name_norm and tag_fam_name_norm in existing_fams:
                    continue

                # Place tag with rotation
                tag = tagger.place_tag_at_center_with_rotation(
                    elem,
                    tag_label=tag_label,
                    position=location_of_tag
                )
                if tag and elem_key is not None and tag_fam_name_norm:
                    if elem_key not in existing_tag_map:
                        existing_tag_map[elem_key] = set()
                    existing_tag_map[elem_key].add(tag_fam_name_norm)

            except Exception as e:
                pass

    t.Commit()
except Exception as e:
    output.print_md("Tag placement error: {}".format(e))
    t.RollBack()
    raise
