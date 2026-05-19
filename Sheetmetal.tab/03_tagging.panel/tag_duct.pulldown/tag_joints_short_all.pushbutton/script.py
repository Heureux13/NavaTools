# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from Autodesk.Revit.DB import Transaction
from pyrevit import revit, forms, DB, script
from ducts.revit_duct import RevitDuct
from tagging.revit_tagging import RevitTagging
from tagging.revit_tagging_joints import Joints
from tagging.tag_config import SLOT_LENGTH, SLOT_STACK
from constants.print_outputs import print_disclaimer

# Button info
# ==================================================
__title__ = "Tag Joints Short All"
__doc__ = """
Tag all short straight duct with length.
Will skip tag if stack is found
"""

# Variables
# ==================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
output = script.get_output()
view = revit.active_view
tagger = RevitTagging(doc=doc, view=view)
joints = Joints(doc=doc, view=view, tagger=tagger)

# Collect ducts in view
# ==================================================
ducts = RevitDuct.all(doc, view)
if not ducts:
    forms.alert("No ducts found in the current view", exitscript=True)

# Slots to check for existing tags (skip if already tagged with these)
skip_if_tagged_with = [SLOT_LENGTH, SLOT_STACK]

# Choose tag
# ==================================================
tag = None
tag_candidates = Joints.TAG_SLOT_CANDIDATES.get(SLOT_LENGTH, [])
for tag_family, slot_type in tag_candidates:
    try:
        tag = tagger.get_label(tag_family)
        break
    except LookupError:
        continue
if tag is None:
    tag_names = [t[0] for t in tag_candidates]
    forms.alert(
        "No tag family found for SLOT_LENGTH. Tried:\n" + "\n".join(tag_names) +
        "\n\nMake sure one of these tag families is loaded in the project.",
        exitscript=True
    )

# Filter ducts
# ==================================================
fil_ducts, skipped_by_param = joints.filter_ducts(ducts)

# Transaction
# ==================================================
already_tagged = []
needs_tagging = []
t = Transaction(doc, "Tag all short duct")
t.Start()
try:
    # Begins our tagging process
    tagged_count = 0
    batch_count = 0
    for d in fil_ducts:
        # Check if already tagged with any of the skip slots
        if joints.is_tagged_with_slots(d.element, skip_if_tagged_with):
            already_tagged.append(d)
            continue

        needs_tagging.append(d)

        # Place tag with rotation
        placed_tag = joints.place_tag_with_rotation(
            d, tag, attempt_rotation=True)
        if placed_tag is not None:
            tagged_count += 1
            batch_count += 1
            if tagged_count % Joints.PROGRESS_EVERY == 0:
                output.print_md("Tagged {} so far...".format(tagged_count))
            if batch_count >= Joints.BATCH_SIZE:
                t.Commit()
                t = Transaction(doc, "Short Joints Tag")
                t.Start()
                batch_count = 0

    # Print newly tagged list first
    if needs_tagging:
        output.print_md("## Newly Tagged")
        for i, d in enumerate(needs_tagging, start=1):
            output.print_md("### No.{} | {}".format(
                i, Joints.format_newly_tagged(d)))
        output.print_md("---")

    # Print already tagged list
    if already_tagged:
        output.print_md("## Already Tagged")
        for i, d in enumerate(already_tagged, start=1):
            output.print_md("### Index {} | {}".format(
                i, Joints.format_already_tagged(d)))
        output.print_md("---")

    # Print skipped by parameter list
    if skipped_by_param:
        output.print_md("## Skipped By Parameter")
        for i, item in enumerate(skipped_by_param, start=1):
            d, skip_name, skip_val = item
            output.print_md("### Index {} | {}".format(
                i, Joints.format_skipped_by_param(d, skip_name, skip_val)))
        output.print_md("---")

    # Summary
    output.print_md("## Summary")
    output.print_md("- **Newly Tagged:** {}".format(len(needs_tagging)))
    output.print_md("- **Already Tagged:** {}".format(len(already_tagged)))
    output.print_md(
        "- **Skipped By Parameter:** {}".format(len(skipped_by_param)))
    output.print_md("- **Total Elements:** {}".format(len(fil_ducts)))
    output.print_md("---")

    # Final helper print
    print_disclaimer(output)

    t.Commit()
except Exception as e:
    output.print_md("Tag placement error: {}".format(e))
    t.RollBack()
    raise
