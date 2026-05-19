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
from ducts.revit_duct import RevitDuct, DuctAngleAllowance
from ducts.revit_xyz import RevitXYZ
from tagging.revit_tagging import RevitTagging
from tagging.revit_tagging_joints import Joints
from tagging.tag_config import SLOT_LENGTH, SLOT_STACK
from constants.print_outputs import print_disclaimer

# Button info
# ==================================================
__title__ = "Tag Joints Short"
__doc__ = """
Tag all short straight duct with length.
"""

# Variables
# ==================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
output = script.get_output()
view = revit.active_view
tagger = RevitTagging(doc=doc, view=view)
joint_tagger = Joints(doc=doc, view=view, tagger=tagger)

# View determination
# ==================================================
if view.ViewType == DB.ViewType.FloorPlan:
    current_view_type = "floor"
elif view.ViewType == DB.ViewType.Section:
    current_view_type = "section"
else:
    current_view_type = "other"

# Custom families for this script (different thresholds than default)
joint_tagger.ELEMENT_FAMILIES = {
    'straight': None,
    'spiral': 6,
    'spiral duct': 6,
}
joint_tagger.DEFAULT_SHORT_THRESHOLD_IN = 56.0

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
# Resolve from shared slot config in tagging classes.
tag_candidates = joint_tagger.TAG_SLOT_CANDIDATES.get(SLOT_LENGTH, [])
for candidate in tag_candidates:
    try:
        if isinstance(candidate, tuple):
            family_name = str(candidate[0]).strip()
            type_name = str(candidate[1]).strip() if len(candidate) > 1 else ''
            if family_name and type_name:
                tag = tagger.get_label_exact(
                    family_name, type_name, allow_fallback=False)
            elif family_name:
                tag = tagger.get_label(family_name)
            else:
                continue
        else:
            tag = tagger.get_label(str(candidate).strip())
        break
    except LookupError:
        continue
if tag is None:
    tag_names = []
    for candidate in tag_candidates:
        if isinstance(candidate, tuple):
            fam = str(candidate[0]).strip()
            typ = str(candidate[1]).strip() if len(candidate) > 1 else ''
            tag_names.append("{}::{}".format(fam, typ) if typ else fam)
        else:
            tag_names.append(str(candidate).strip())
    forms.alert(
        "No tag family found for SLOT_LENGTH. Tried:\n" + "\n".join(tag_names) +
        "\n\nMake sure one of these tag families is loaded in the project.",
        exitscript=True
    )

# Filter ducts with base filtering
# ==================================================
fil_ducts_base, skipped_by_param = joint_tagger.filter_ducts(ducts)

# Additional filtering: spiral length and angle-based filtering
# ==================================================
fil_ducts = []
for d in fil_ducts_base:
    fam = (d.family or "").strip().lower()

    # Spiral parts should be considered short up to 10'-0" regardless of
    # connector metadata quality; this avoids false negatives in joint_size.
    if fam in ('spiral', 'spiral duct'):
        if d.length is None or d.length > 120.0:
            continue

    # Angle-based filtering based on view type
    angle = RevitXYZ(d.element).straight_joint_degree()
    if isinstance(angle, (int, float)):
        abs_angle = abs(angle)
        if current_view_type == "floor":
            if DuctAngleAllowance.VERTICAL.contains(abs_angle):
                continue
        elif current_view_type == "section":
            if DuctAngleAllowance.HORIZONTAL.contains(abs_angle):
                continue

    fil_ducts.append(d)

# Transaction
# ==================================================
already_tagged = []
newly_tagged = []
could_not_place = []
t = Transaction(doc, "Short Joints Tag")
t.Start()
try:
    # Begins our tagging process
    tagged_count = 0
    batch_count = 0
    for d in fil_ducts:
        # Check if already tagged with any of the skip slots
        if joint_tagger.is_tagged_with_slots(d.element, skip_if_tagged_with):
            already_tagged.append(d)
            continue

        # Place tag (without rotation for this script)
        placed_tag = joint_tagger.place_tag_with_rotation(
            d, tag, attempt_rotation=False)
        if placed_tag is not None:
            newly_tagged.append(d)
            tagged_count += 1
            batch_count += 1
            if tagged_count % Joints.PROGRESS_EVERY == 0:
                output.print_md("Tagged {} so far...".format(tagged_count))
            if batch_count >= Joints.BATCH_SIZE:
                t.Commit()
                t = Transaction(doc, "Short Joints Tag")
                t.Start()
                batch_count = 0
        else:
            could_not_place.append(d)

    # Print newly tagged list first
    if newly_tagged:
        output.print_md("## Newly Tagged")
        for i, d in enumerate(newly_tagged, start=1):
            output.print_md(
                "### No.{} | ID: {} | Fam: {} | Size: {} | Le: {:06.2f} | Ex: {}".format(
                    i,
                    output.linkify(d.element.Id),
                    d.family,
                    d.size,
                    d.length if d.length else 0.0,
                    d.extension_bottom if d.extension_bottom else 0.0
                )
            )
        output.print_md("---")

    if could_not_place:
        output.print_md("## Could Not Place Tag")
        for i, d in enumerate(could_not_place, start=1):
            output.print_md(
                "### Index {} | Family: {} | Length: {:06.2f} | Element ID: {}".format(
                    i,
                    d.family,
                    d.length if d.length else 0.0,
                    output.linkify(d.element.Id)
                )
            )
        output.print_md("---")

    # Print already tagged list
    if already_tagged:
        output.print_md("## Already Tagged")
        for i, d in enumerate(already_tagged, start=1):
            output.print_md(
                "### Index {} | Size: {} | Family: {} | Length: {:06.2f} | Element ID: {}".format(
                    i,
                    d.size,
                    d.family,
                    d.length if d.length else 0.0,
                    output.linkify(d.element.Id)
                )
            )
        output.print_md("---")

    # Print skipped by parameter list
    if skipped_by_param:
        output.print_md("## Skipped By Parameter")
        for i, item in enumerate(skipped_by_param, start=1):
            d, skip_name, skip_val = item
            output.print_md(
                "### Index {} | Param: {} | Value: {} | Family: {} | Length: {:06.2f} | Element ID: {}".format(
                    i,
                    skip_name,
                    skip_val,
                    d.family,
                    d.length if d.length else 0.0,
                    output.linkify(d.element.Id)
                )
            )
        output.print_md("---")

    # Summary
    output.print_md("## Summary")
    output.print_md("- **Newly Tagged:** {}".format(len(newly_tagged)))
    output.print_md("- **Already Tagged:** {}".format(len(already_tagged)))
    output.print_md(
        "- **Skipped By Parameter:** {}".format(len(skipped_by_param)))
    output.print_md(
        "- **Could Not Place Tag:** {}".format(len(could_not_place)))
    output.print_md("- **Total Elements:** {}".format(len(fil_ducts)))
    output.print_md("---")

    # Final helper print
    print_disclaimer(output)

    t.Commit()
except Exception as e:
    output.print_md("Tag placement error: {}".format(e))
    t.RollBack()
    raise
