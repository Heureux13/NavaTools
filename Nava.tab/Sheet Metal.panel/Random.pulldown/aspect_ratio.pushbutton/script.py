# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from revit_duct import RevitDuct
from size import Size
from Autodesk.Revit.DB import Transaction
from pyrevit import revit, script

# Button display information
# =================================================
__title__ = "Set Aspect Ratio"
__doc__ = """
Gives a rounded duct ratio
"""


# Code
# ==================================================
doc = revit.doc
view = revit.active_view
output = script.get_output()

uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

ducts = RevitDuct.all(doc, view)

# Filter for straight joints only
straight_joint_families = {
    'straight',
    'spiral tube',
    'round duct',
    'tube',
    'spiral duct',
}

all_straights = [
    d for d in ducts
    if d.family and d.family.strip().lower() in straight_joint_families
]

output.print_md("## Found {} straight joints".format(len(all_straights)))

if not all_straights:
    output.print_md("No straight joints found in view.")
    script.exit()


# Helper functions
# ==================================================
def calculate_aspect_ratio(size_str):
    try:
        size = Size(size_str)

        # Get width and height
        width = size.in_width
        height = size.in_height

        if width is None or height is None:
            return None

        # Always use larger dimension as numerator to keep ratio >= 1:1
        larger = max(width, height)
        smaller = min(width, height)

        if smaller == 0:
            return None

        ratio = larger / smaller

        # Round to nearest tenth
        ratio_rounded = round(ratio, 1)

        return "{}:1".format(ratio_rounded)
    except Exception as e:
        return None


# Main transaction
# ==================================================
t = Transaction(doc, "Set Aspect Ratios")
t.Start()

try:
    successful = []
    failed = []
    skipped = []

    for duct in all_straights:
        try:
            # Get size parameter
            size_param = duct.element.LookupParameter("Size")
            if not size_param or not size_param.AsString():
                skipped.append((duct, "No Size parameter"))
                continue

            size_str = size_param.AsString()

            # Calculate aspect ratio
            aspect_ratio = calculate_aspect_ratio(size_str)
            if aspect_ratio is None:
                skipped.append(
                    (duct, "Could not calculate ratio from size '{}'".format(size_str)))
                continue

            # Set the aspect ratio parameter
            aspect_param = duct.element.LookupParameter("_duct_aspect_ratio")
            if not aspect_param:
                skipped.append(
                    (duct, "_duct_aspect_ratio parameter not found"))
                continue

            aspect_param.Set(aspect_ratio)
            successful.append((duct, size_str, aspect_ratio))

        except Exception as e:
            failed.append((duct, str(e)))

    t.Commit()

    # Output results
    # ==================================================
    output.print_md("## Aspect Ratio Assignment Results")
    output.print_md("---")
    output.print_md("### Successful: {} elements".format(len(successful)))

    if successful:
        for duct, size, ratio in successful:
            output.print_md(
                "- ID: {} | Family: {} | Size: {} → Aspect Ratio: {}".format(
                    output.linkify(duct.element.Id),
                    duct.family,
                    size,
                    ratio
                ))

    output.print_md("---")

    if failed:
        output.print_md("### Failed: {} elements".format(len(failed)))
        for duct, reason in failed:
            output.print_md(
                "- ID: {} | Reason: {}".format(
                    output.linkify(duct.element.Id),
                    reason
                ))
        output.print_md("---")

    if skipped:
        output.print_md("### Skipped: {} elements".format(len(skipped)))
        for duct, reason in skipped:
            output.print_md(
                "- ID: {} | Family: {} | Size: {} | Reason: {}".format(
                    output.linkify(duct.element.Id),
                    duct.family,
                    duct.size if duct.size else "N/A",
                    reason
                ))

    output.print_md("---")
    output.print_md("## Summary")
    output.print_md("- **Successful**: {}".format(len(successful)))
    output.print_md("- **Failed**: {}".format(len(failed)))
    output.print_md("- **Skipped**: {}".format(len(skipped)))
    output.print_md(
        "- **Total Processed**: {}".format(len(successful) + len(failed) + len(skipped)))

except Exception as e:
    t.RollBack()
    output.print_md("## Error: {}".format(str(e)))
    raise
