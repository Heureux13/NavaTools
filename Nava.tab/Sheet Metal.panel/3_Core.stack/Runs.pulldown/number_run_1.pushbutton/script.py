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
from revit_element import RevitElement
from revit_output import print_disclaimer
from revit_runs import RrevitRuns
from pyrevit import revit, script
from Autodesk.Revit.DB import *

# Button info
# ===================================================
__title__ = "Numbers Run w/o Matching"
__doc__ = """
1, 2, 3... n. No matches will be considered.
"""

# Variables
# ==================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
view = revit.active_view
output = script.get_output()

# Main Code
# ==================================================
# NUMBERING ALGORITHM:
# 1. User selects a fitting (e.g., fitting C with value 15)
# 2. Check if selected fitting is in number_families and doesn't have skip_value
# 3. Find the next sequential number in connected fittings:
#    - Look for connected fitting with value 16 (current + 1)
#    - From 16, look for connected fitting with value 17
#    - Continue following the number chain until next number not found
# 4. Once chain is complete, number remaining unvisited connections sequentially
# 5. This ensures directional flow without backtracking

# Parameters that hold the item number (will be matched case-insensitive)
number_paramters = {
    'item number',
    '_umi_item_number',
}

match_paramters = {
    'family',
    'size',
    'length',
    'angle',
}

# Families allowed to be numbered
number_families = {
    "straight",
    "transition",
    "radius elbow",
    "elbow - 90 degree",
    "elbow",
    "drop cheek",
    "ogee",
    "offset",
    "square to ø",
    "end cap",
    "tdf end cap",
    'reducer',
    'tee',
}

# Families not allowed to be numbered but allowed to traverse through
allow_but_not_number = {
    'manbars',
    'canvas',
    'fire damper - type a',
    'fire damper - type b',
    'fire damper - type c',
    'fire damper - type cr',
    'smoke fire damper - type cr',
    'smoke fire damper - type csr',
    'rect volume damper',
    'access door',
    "straight tap"
}

# Values that indicate to traverse through but not number
skip_values = {
    0,
    "skip",
    "n/a",
}

# Values that indicate to stop the run (do not traverse beyond)
stop_values = {
    "stop",
}

# Families that need to be numbered after their connected run has been numbered
store_families = {
    'boot tap',
    'straight tap',
    'rec on rnd straight tap',
}

# Helper Functions
# ==================================================

runs = RrevitRuns(
    doc=doc,
    view=view,
    number_paramters=number_paramters,
    skip_values=skip_values,
    stop_values=stop_values,
    number_families=number_families,
    allow_but_not_number=allow_but_not_number,
    store_families=store_families,
)


# Main Script
# ==================================================
# Get selected fitting
selected_duct = RevitDuct.from_selection(uidoc, doc, view)
selected_duct = selected_duct[0] if selected_duct else None


# Start of numbering logic
if selected_duct:
    # Validate selected fitting - allow store_families when selected
    selected_family = selected_duct.family
    selected_family_lower = selected_family.lower() if selected_family else ""
    is_store_family = selected_family_lower in store_families

    if not runs.is_numberable(selected_duct) and not is_store_family:
        output.print_md("## Selected fitting '{}' cannot be numbered".format(
            selected_duct.family if selected_duct.family else "Unknown"
        ))
    elif runs.has_skip_value(selected_duct):
        output.print_md(
            "## Selected fitting has a skip value and cannot be numbered")
    else:
        # Start transaction
        t = Transaction(doc, "Number Duct Run")
        t.Start()

        try:
            modified_ducts = []

            start_number = runs.get_item_number(selected_duct)
            if start_number is None:
                start_number = 1

            visited = {selected_duct.id}
            runs.set_item_number(selected_duct, start_number)
            modified_ducts.append(selected_duct)

            stored_taps = []
            allow_stores = is_store_family

            filter_size = None
            if is_store_family and selected_duct.size_out:
                filter_size = selected_duct.size_out

            last_number, stored_taps, forward_modified, forward_count = runs.number_run_forward(
                selected_duct,
                start_number + 1,
                visited,
                stored_taps,
                [],
                allow_stores,
                filter_size
            )
            modified_ducts.extend(forward_modified)

            if stored_taps:
                branches_to_process = [tap_duct for tap_duct, _ in stored_taps]

                while branches_to_process:
                    branch_duct = branches_to_process.pop(0)

                    if branch_duct.id in visited and not (
                        branch_duct.family and branch_duct.family.lower() in store_families
                    ):
                        continue

                    anchor_num, anchor_duct = runs.find_connected_numbered_element(branch_duct)

                    if anchor_num is None:
                        continue

                    base_for_branch = (last_number + 1) if last_number is not None else (anchor_num + 1)
                    branch_start = runs.round_up_to_nearest_10(base_for_branch)

                    filter_size = branch_duct.size_out if branch_duct.family and branch_duct.family.lower() in store_families else None

                    sub_branches = []

                    if not runs.has_skip_value(branch_duct):
                        runs.set_item_number(branch_duct, branch_start)
                        modified_ducts.append(branch_duct)

                    branch_first = branch_start + 1
                    branch_last = runs.number_branch_recursive(
                        branch_duct,
                        branch_first,
                        visited,
                        sub_branches,
                        modified_ducts,
                        filter_size,
                        skip_start_numbering=True
                    )

                    if branch_last > last_number:
                        last_number = branch_last

                    if sub_branches:
                        branches_to_process = sub_branches + branches_to_process

            output.print_md(
                "# Total elements: {:03d}, {}".format(
                    len(modified_ducts),
                    output.linkify([d.element.Id for d in modified_ducts])
                ))

            if modified_ducts:
                output.print_md("\n## Detailed Element Information:")
                for duct in modified_ducts:
                    item_num = runs.get_item_number(duct)
                    family = duct.family if duct.family else "Unknown"

                    # Get connected elements
                    connected = runs.get_connected_fittings(duct)
                    connected_info = []
                    for conn in connected:
                        conn_family = conn.family if conn.family else "Unknown"
                        conn_num = runs.get_item_number(conn)
                        connected_info.append("{}[#{}]".format(conn_family, conn_num if conn_num else "None"))

                    connected_str = ", ".join(connected_info) if connected_info else "None"

                    output.print_md(
                        "- {} | Item#: **{}** | Family: **{}** | Connected: {}".format(
                            output.linkify(duct.element.Id),
                            item_num if item_num else "None",
                            family,
                            connected_str
                        )
                    )

                try:
                    nums = [runs.get_item_number(d) for d in modified_ducts]
                    nums = [n for n in nums if n is not None]
                    if nums:
                        output.print_md("\n**Start: {} | End: {}**".format(min(nums), max(nums)))
                except Exception:
                    pass

            RevitElement.select_many(uidoc, modified_ducts)
            t.Commit()

        except Exception as e:
            t.RollBack()
            output.print_md("## Error during numbering: {}".format(str(e)))
            import traceback
            output.print_md("```\n{}\n```".format(traceback.format_exc()))

    # Final print statements
    print_disclaimer(output)
else:
    output.print_md("## Select a fitting first")
