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
from pyrevit import revit, script
from Autodesk.Revit.DB import *
import re

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
    'conical tee',
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
}

# Values that indicate to traverse through but not number
skip_values = {
    0,
    "skip",
    "n/a",
}

# Families that need to be numbered after their connected run has been numbered
store_families = {
    'boot tap',
    'straight tap',
    'rec on rnd straight tap',
}

# Helper Functions
# ==================================================


def get_item_number(duct):
    """Get the current item number from any of the number parameters."""
    # Get all parameters and search case-insensitively
    for param in duct.element.Parameters:
        param_name_lower = param.Definition.Name.strip().lower()
        if param_name_lower not in number_paramters:
            continue

        # Get the value as string
        val = param.AsString()
        if val is None:
            val = param.AsValueString()

        if val is not None:
            # Try to convert to int, if it's a skip value return None
            try:
                num_val = int(val) if isinstance(
                    val, (int, float)) else int(float(val))
                if num_val in skip_values:
                    return None
                return num_val
            except (ValueError, TypeError):
                # Check if string value is in skip_values
                if val in skip_values:
                    return None
                # Try to extract number from string
                match = re.search(r'\d+', str(val))
                if match:
                    return int(match.group())
    return None


def set_item_number(duct, number):
    """Set the item number in the first available parameter."""
    # Get all parameters and search case-insensitively
    for param in duct.element.Parameters:
        param_name = param.Definition.Name
        param_name_lower = param_name.strip().lower()

        if param_name_lower not in number_paramters:
            continue

        if param.IsReadOnly:
            output.print_md(
                "  *Warning: Parameter '{}' is read-only*".format(param_name))
            continue

        # Try different approaches based on storage type
        try:
            storage_type = param.StorageType
            output.print_md("  *Debug: Found parameter '{}', storage type: {}*".format(
                param_name, storage_type))

            if storage_type == StorageType.String:
                param.Set(str(number))
                output.print_md(
                    "  *Success: Set '{}' to '{}'*".format(param_name, number))
                return True
            elif storage_type == StorageType.Integer:
                param.Set(int(number))
                output.print_md(
                    "  *Success: Set '{}' to {}*".format(param_name, number))
                return True
            elif storage_type == StorageType.Double:
                param.Set(float(number))
                output.print_md(
                    "  *Success: Set '{}' to {}*".format(param_name, number))
                return True
            else:
                output.print_md("  *Warning: Parameter '{}' has unsupported storage type: {}*".format(
                    param_name, storage_type))
        except Exception as e:
            output.print_md(
                "  *Error setting parameter '{}': {}*".format(param_name, str(e)))
            continue

    output.print_md(
        "  *Warning: Could not set item number - no writable parameter found*")
    return False


def get_connected_fittings(duct, doc, view):
    """Get all immediately connected fittings (only direct connections)."""
    connected = []
    for connector in duct.get_connectors():
        if not connector.IsConnected:
            continue
        all_refs = list(connector.AllRefs)
        for ref in all_refs:
            if ref and hasattr(ref, 'Owner'):
                connected_elem = ref.Owner
                # Only process fabrication parts
                if not isinstance(connected_elem, FabricationPart):
                    continue
                # Skip the same element
                if connected_elem.Id == duct.element.Id:
                    continue
                try:
                    connected_duct = RevitDuct(doc, view, connected_elem)
                    connected.append(connected_duct)
                except Exception:
                    continue
    return connected


def is_numberable(duct):
    """Check if a duct can be numbered based on family."""
    family = duct.family
    if not family:
        return False
    family_lower = family.lower()
    return family_lower in number_families


def is_traversable(duct):
    """Check if we can traverse through this duct (even if not numbering it)."""
    family = duct.family
    if not family:
        return False
    family_lower = family.lower()
    return family_lower in allow_but_not_number or is_numberable(duct)


def has_skip_value(duct):
    """Check if duct has a skip value in its number parameter."""
    # Get all parameters and search case-insensitively
    for param in duct.element.Parameters:
        param_name_lower = param.Definition.Name.strip().lower()
        if param_name_lower not in number_paramters:
            continue

        val = param.AsString()
        if val is None:
            val = param.AsValueString()

        if val in skip_values:
            return True
        try:
            if int(val) in skip_values:
                return True
        except (ValueError, TypeError):
            pass
    return False


def get_match_signature(duct):
    """
    Get the match signature for a duct based on match_parameters.
    Returns a tuple of (family, size, length, angle) for comparison.
    """
    signature = []

    # Family
    family = duct.family if duct.family else ""
    signature.append(family.lower())

    # Size
    size = duct.size if hasattr(duct, 'size') and duct.size else ""
    signature.append(str(size))

    # Length - get from parameter
    length = ""
    try:
        for param in duct.element.Parameters:
            param_name_lower = param.Definition.Name.strip().lower()
            if param_name_lower == 'length':
                val = param.AsString()
                if val is None:
                    val = param.AsValueString()
                if val:
                    length = str(val)
                break
    except Exception:
        pass
    signature.append(length)

    # Angle - get from parameter
    angle = ""
    try:
        for param in duct.element.Parameters:
            param_name_lower = param.Definition.Name.strip().lower()
            if param_name_lower == 'angle':
                val = param.AsString()
                if val is None:
                    val = param.AsValueString()
                if val:
                    angle = str(val)
                break
    except Exception:
        pass
    signature.append(angle)

    return tuple(signature)


def find_matching_elements(numbered_ducts, signature_map, doc, view):
    """
    Find all elements in the model that match the numbered ducts based on match_parameters.
    Uses the existing signature_map from numbering process.
    Returns a dictionary: {signature: [(duct, item_number), ...]}
    """
    # Get all fabrication parts in the model
    collector = FilteredElementCollector(doc, view.Id)
    all_fab_parts = collector.OfClass(FabricationPart).ToElements()

    # Find matches
    matches = {}
    for elem in all_fab_parts:
        try:
            candidate_duct = RevitDuct(doc, view, elem)

            # Skip if already numbered in this run
            if candidate_duct.id in [d.id for d in numbered_ducts]:
                continue

            # Skip if not numberable
            if not is_numberable(candidate_duct):
                continue

            # Check if it matches any signature
            candidate_sig = get_match_signature(candidate_duct)
            if candidate_sig in signature_map:
                if candidate_sig not in matches:
                    matches[candidate_sig] = []
                # Use the item number from the signature map
                item_num = signature_map[candidate_sig]
                matches[candidate_sig].append((candidate_duct, item_num))
        except Exception:
            continue

    return matches


def find_duct_with_number(connected_ducts, target_number):
    """
    Find a connected fitting with a specific number.
    Returns the duct with that number or None if not found.
    """
    for duct in connected_ducts:
        num = get_item_number(duct)
        if num == target_number:
            return duct
    return None


def follow_number_chain(start_duct, doc, view, visited=None):
    """
    Follow the existing number chain from the start fitting.
    Returns (last_duct_in_chain, last_number_in_chain, visited_in_chain, chain_ducts).
    """
    if visited is None:
        visited = set()

    chain_ducts = []
    current_duct = start_duct
    current_number = get_item_number(current_duct)

    if current_number is None:
        # No number on start fitting
        return (current_duct, None, visited, chain_ducts)

    visited.add(current_duct.id)
    chain_ducts.append(current_duct)

    # Follow the chain forward by looking for the next sequential number
    while True:
        next_number = current_number + 1
        connected = get_connected_fittings(current_duct, doc, view)

        # Filter to unvisited and traversable
        unvisited_traversable = [
            c for c in connected
            if c.id not in visited and is_traversable(c)
        ]

        # Look for the next number in connected fittings
        next_duct = find_duct_with_number(unvisited_traversable, next_number)

        if next_duct is None:
            # Chain ends here
            break

        visited.add(next_duct.id)
        chain_ducts.append(next_duct)
        current_duct = next_duct
        current_number = next_number
        output.print_md("*Followed chain: Found {} with number {}*".format(
            current_duct.family if current_duct.family else "Unknown",
            current_number
        ))

    return (current_duct, current_number, visited, chain_ducts)


def find_endpoints(start_duct, doc, view, visited=None):
    """
    Find all fittings in the run that are true endpoints (only 1 traversable connection total).
    Returns a list of duct objects that are endpoints.
    """
    if visited is None:
        visited = set()

    endpoints = []
    all_ducts = []
    to_process = [start_duct]

    # First, collect all traversable ducts in the run
    while to_process:
        duct = to_process.pop(0)

        if duct.id in visited:
            continue
        visited.add(duct.id)
        all_ducts.append(duct)

        # Get all connected fittings
        connected = get_connected_fittings(duct, doc, view)
        for conn in connected:
            if conn.id not in visited and is_traversable(conn):
                to_process.append(conn)

    # Now find true endpoints (ducts with only 1 traversable connection)
    for duct in all_ducts:
        connected = get_connected_fittings(duct, doc, view)
        traversable_count = sum(1 for c in connected if is_traversable(c))

        if traversable_count == 1:
            endpoints.append(duct)
            output.print_md("*Found endpoint: {} (ID: {})*".format(
                duct.family if duct.family else "Unknown",
                output.linkify(duct.element.Id)
            ))

    return endpoints


def find_anchor_number(duct, doc, view, visited=None):
    """
    Recursively search backwards through connections to find an existing number.
    Returns (anchor_number, anchor_duct) or (None, None) if no anchor found.
    """
    if visited is None:
        visited = set()

    visited.add(duct.id)

    # Check if this duct has a number
    current_number = get_item_number(duct)
    if current_number is not None and current_number > 0:
        return (current_number, duct)

    # Get connected fittings and search through them
    connected = get_connected_fittings(duct, doc, view)
    for conn in connected:
        if conn.id in visited:
            continue

        # Only traverse through valid families
        if not is_traversable(conn):
            continue

        # Recursively search
        anchor_num, anchor_duct = find_anchor_number(conn, doc, view, visited)
        if anchor_num is not None:
            return (anchor_num, anchor_duct)

    return (None, None)


def number_run_forward(
        start_duct,
        start_number,
        doc,
        view,
        visited=None,
        stored_taps=None,
        modified_ducts=None,
        allow_store_families=False):
    """
    Number fittings sequentially starting from start_duct with start_number.
    Simply increments the number for each numberable fitting (no duplicate matching).
    allow_store_families: If True, store_families can be numbered (used when they are selected)
    Returns the last number used, list of stored tap fittings, and modified ducts.
    """
    if visited is None:
        visited = set()
    if stored_taps is None:
        stored_taps = []
    if modified_ducts is None:
        modified_ducts = []

    current_number = start_number

    # Get connections from the start duct
    connected = get_connected_fittings(start_duct, doc, view)
    output.print_md("*Found {} connected fittings*".format(len(connected)))

    to_process = [(conn, current_number)
                  for conn in connected if conn.id not in visited]

    while to_process:
        duct, num = to_process.pop(0)

        if duct.id in visited:
            continue
        visited.add(duct.id)

        # Check if this is a store_family (tap)
        family = duct.family
        family_lower = family.lower() if family else ""

        output.print_md("*Checking: {} (family: {})*".format(
            output.linkify(duct.element.Id),
            family if family else "None"
        ))

        if family_lower in store_families:
            if not allow_store_families:
                stored_taps.append((duct, None))
                output.print_md("  *Stored tap for later (no value written)*")
                # Skip during traversal unless it was the selected fitting
                continue
            # If allow_store_families is True, fall through to number it

        # Check if we should number this fitting
        if is_numberable(duct):
            if has_skip_value(duct):
                output.print_md("  *Has skip value, writing 'skip'*")
                # Has skip value - write "skip" to mark it
                set_item_number(duct, "skip")
                modified_ducts.append(duct)
            else:
                # Simply assign the next sequential number
                set_item_number(duct, current_number)
                output.print_md("Set {} to **{}** (ID: {})".format(
                    family if family else "Unknown",
                    current_number,
                    output.linkify(duct.element.Id)
                ))
                modified_ducts.append(duct)
                current_number += 1
        elif is_traversable(duct):
            # Don't number but continue traversing
            output.print_md(
                "  *Traversable (allow_but_not_number), continuing*")
            pass
        else:
            # Can't traverse through this
            output.print_md("  *Not numberable or traversable, skipping*")
            continue

        # Get next connections
        next_connected = get_connected_fittings(duct, doc, view)
        for conn in next_connected:
            if conn.id not in visited:
                to_process.append((conn, current_number))

    return current_number - 1, stored_taps, modified_ducts, len(modified_ducts)


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

    if not is_numberable(selected_duct) and not is_store_family:
        output.print_md("## Selected fitting '{}' cannot be numbered".format(
            selected_duct.family if selected_duct.family else "Unknown"
        ))
    elif has_skip_value(selected_duct):
        output.print_md(
            "## Selected fitting has a skip value and cannot be numbered")
    else:
        # Start transaction
        t = Transaction(doc, "Number Duct Run")
        t.Start()

        try:
            output.print_md("## Starting numbering process...")
            output.print_md("---")

            # Track all modified ducts
            modified_ducts = []

            # Check if selected fitting is a store_family
            is_selected_store_family = is_store_family

            # Get the starting number from the selected fitting's parameter
            start_number = get_item_number(selected_duct)
            if start_number is None:
                start_number = 1

            output.print_md("Starting at **{}**".format(start_number))
            output.print_md("---")

            visited = {selected_duct.id}

            # Number the selected fitting first (use the parameter value or 1)
            set_item_number(selected_duct, start_number)
            modified_ducts.append(selected_duct)
            output.print_md("Set {} to **{}** (ID: {})".format(
                selected_duct.family if selected_duct.family else "Unknown", a
                start_number,
                output.linkify(selected_duct.element.Id)
            ))

            output.print_md("---")

            # Continue numbering forward from selected fitting
            stored_taps = []
            allow_stores = is_selected_store_family if 'is_selected_store_family' in locals() else False
            last_number, stored_taps, forward_modified, forward_count = number_run_forward(
                selected_duct,
                start_number + 1,
                doc,
                view,
                visited,
                stored_taps,
                [],
                allow_stores
            )
            modified_ducts.extend(forward_modified)
            modified_count = 1 + forward_count

            # Store tap fittings for future processing (not numbering them yet)
            if stored_taps:
                output.print_md("---")
                output.print_md(
                    "## Stored {} tap fittings for future processing (not numbered)".format(len(stored_taps)))
                for tap_duct, tap_num in stored_taps:
                    output.print_md("*Stored: {} (ID: {})*".format(
                        tap_duct.family if tap_duct.family else "Unknown",
                        output.linkify(tap_duct.element.Id)
                    ))

            output.print_md("---")
            output.print_md(
                "## Numbering complete! Last number used: **{}** | Total modified: **{}**".format(last_number, modified_count))

            # Select all modified ducts
            if modified_ducts:
                RevitElement.select_many(uidoc, modified_ducts)
                output.print_md(
                    "## Selected {} total modified fittings".format(modified_count))

            # Commit transaction
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
