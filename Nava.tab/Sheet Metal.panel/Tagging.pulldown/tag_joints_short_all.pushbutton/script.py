# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
import math
from Autodesk.Revit.DB import Transaction
from pyrevit import revit, forms, DB, script
from revit_duct import RevitDuct, JointSize, DuctAngleAllowance
from revit_xyz import RevitXYZ
from revit_tagging import RevitTagging
from revit_output import print_disclaimer

# Button info
# ==================================================
__title__ = "Tag All Joints Short"
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
DEFAULT_SHORT_THRESHOLD_IN = 56.0
PROGRESS_EVERY = 500
BATCH_SIZE = 200

# View determination
# ==================================================
if view.ViewType == DB.ViewType.FloorPlan:
    current_view_type = "floor"
elif view.ViewType == DB.ViewType.Section:
    current_view_type = "section"
else:
    current_view_type = "other"

# Collect ducts in view
# ==================================================
ducts = RevitDuct.all(doc, view)
if not ducts:
    forms.alert("No ducts found in the current view", exitscript=True)

element_families = {
    'straight': None,
    'spiral': 12,
    'spiral duct': 12,
}

skip_parameters = {
    'mark': ['skip', 'skip n/a'],
}


def should_skip_by_param(element, param_rules):
    for param_name, skip_values in param_rules.items():
        param = element.LookupParameter(param_name)
        if not param:
            continue
        raw_val = None
        try:
            raw_val = param.AsString()
        except Exception:
            raw_val = None
        if not raw_val:
            try:
                raw_val = param.AsValueString()
            except Exception:
                raw_val = None
        if raw_val is None:
            continue
        val = raw_val.strip().lower()
        if val in {v.strip().lower() for v in skip_values}:
            return True, param_name, raw_val
    return False, None, None


# Choose tag
# ==================================================
tag = tagger.get_label("-FabDuct_LENGTH_FIX_Tag")

# Filtered results
# ==================================================
fil_ducts = []
skipped_by_param = []
for d in ducts:
    # Check if family matches allowed families
    fam = (d.family or "").strip().lower()
    if fam not in element_families:
        continue

    # Skip when parameter exists and matches skip list
    skip_param, skip_name, skip_val = should_skip_by_param(d.element, skip_parameters)
    if skip_param:
        skipped_by_param.append((d, skip_name, skip_val))
        continue

    # Check minimum length threshold for this family
    min_length = element_families.get(fam)
    if min_length is not None:
        length_val = d.length
        # Handle different length types (float, int, or string)
        if length_val is not None:
            try:
                if isinstance(length_val, str):
                    length_val = float(length_val)
                if isinstance(length_val, (int, float)) and length_val <= min_length:
                    continue
            except (ValueError, TypeError):
                pass

    joint_size = d.joint_size
    if joint_size == JointSize.INVALID:
        if d.length is None or d.length > DEFAULT_SHORT_THRESHOLD_IN:
            continue
    elif joint_size != JointSize.SHORT:
        continue
    fil_ducts.append(d)

# Transaction
# ==================================================
already_tagged = []
needs_tagging = []
t = Transaction(doc, "Short Joints Tag")
t.Start()
try:
    # Get tag family name once
    tag_fam_name = (tag.Family.Name if tag and tag.Family else "").strip().lower()

    # Begins our tagging process
    tagged_count = 0
    batch_count = 0
    for d in fil_ducts:
        # Get existing tag families for this element
        existing_tag_fams = tagger.get_existing_tag_families(d.element)

        # Check if already tagged with this tag family
        if tag_fam_name in existing_tag_fams:
            already_tagged.append(d)
            continue

        needs_tagging.append(d)

        # Get the angle for rotation
        angle_deg = None
        try:
            angle_deg = RevitXYZ(d.element).straight_joint_degree()
        except Exception:
            angle_deg = None

        # Convert degrees to radians for Revit API
        angle_rad = None
        if isinstance(angle_deg, (int, float)):
            import math
            angle_rad = math.radians(angle_deg)

        # Place tag and set rotation
        placed_tag = None
        loc = d.element.Location
        if hasattr(loc, "Point") and loc.Point is not None:
            placed_tag = tagger.place_tag(d.element, tag, loc.Point)
            tagged_count += 1
            batch_count += 1
            if angle_rad is not None and placed_tag is not None:
                try:
                    placed_tag.Rotation = angle_rad
                except Exception:
                    pass
            if tagged_count % PROGRESS_EVERY == 0:
                output.print_md("Tagged {} so far...".format(tagged_count))
            if batch_count >= BATCH_SIZE:
                t.Commit()
                t = Transaction(doc, "Short Joints Tag")
                t.Start()
                batch_count = 0
            continue
        if hasattr(loc, "Curve") and loc.Curve is not None:
            curve = loc.Curve
            midpoint = curve.Evaluate(0.5, True)
            placed_tag = tagger.place_tag(d.element, tag, midpoint)
            tagged_count += 1
            batch_count += 1
            if angle_rad is not None and placed_tag is not None:
                try:
                    placed_tag.Rotation = angle_rad
                except Exception:
                    pass
            if tagged_count % PROGRESS_EVERY == 0:
                output.print_md("Tagged {} so far...".format(tagged_count))
            if batch_count >= BATCH_SIZE:
                t.Commit()
                t = Transaction(doc, "Short Joints Tag")
                t.Start()
                batch_count = 0
            continue

        ref, centroid = tagger.get_face_facing_view(d.element)
        if ref is not None and centroid is not None:
            placed_tag = tagger.place_tag(ref, tag, centroid)
            tagged_count += 1
            batch_count += 1
            if angle_rad is not None and placed_tag is not None:
                try:
                    placed_tag.Rotation = angle_rad
                except Exception:
                    pass
            if tagged_count % PROGRESS_EVERY == 0:
                output.print_md("Tagged {} so far...".format(tagged_count))
            if batch_count >= BATCH_SIZE:
                t.Commit()
                t = Transaction(doc, "Short Joints Tag")
                t.Start()
                batch_count = 0
            continue

    # Print newly tagged list first
    if needs_tagging:
        output.print_md("## Newly Tagged")
        for i, d in enumerate(needs_tagging, start=1):
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
    output.print_md("- **Newly Tagged:** {}".format(len(needs_tagging)))
    output.print_md("- **Already Tagged:** {}".format(len(already_tagged)))
    output.print_md("- **Skipped By Parameter:** {}".format(len(skipped_by_param)))
    output.print_md("- **Total Elements:** {}".format(len(fil_ducts)))
    output.print_md("---")

    # Final helper print
    print_disclaimer(output)

    t.Commit()
except Exception as e:
    output.print_md("Tag placement error: {}".format(e))
    t.RollBack()
    raise
