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
from size import Size
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


def round_up_to_nearest_10(number):
    """Round up to the nearest 10. E.g., 55 -> 60, 60 -> 60, 1 -> 10"""
    import math
    return int(math.ceil(number / 10.0) * 10)


def _size_signature(size_value):
    """Create a normalized size signature for matching sizes."""
    if size_value is None:
        return None

    size_str = str(size_value).strip()
    if not size_str:
        return None

    size_obj = Size(size_str)

    # Round
    if size_obj.in_diameter is not None:
        return ("round", round(float(size_obj.in_diameter), 4))

    # Oval
    if size_obj.in_oval_dia is not None:
        w = size_obj.in_width
        h = size_obj.in_height
        if w is not None and h is not None:
            return ("oval", round(float(w), 4), round(float(h), 4))

    # Rectangle / square (order-independent)
    if size_obj.in_width is not None and size_obj.in_height is not None:
        dims = sorted([round(float(size_obj.in_width), 4), round(float(size_obj.in_height), 4)])
        return ("rect", tuple(dims))

    return None


def is_rectangular_size(size_value):
    """Check if a size is rectangular (not round or oval)."""
    sig = _size_signature(size_value)
    return sig is not None and sig[0] == "rect"


def sizes_match(filter_size, conn_size):
    """Return True if sizes match, ignoring quotes and width/height order."""
    sig_a = _size_signature(filter_size)
    sig_b = _size_signature(conn_size)
    if sig_a is None or sig_b is None:
        return False
    return sig_a == sig_b


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
            # Check if it's a skip value first (case-insensitive)
            val_lower = str(val).strip().lower()
            if val_lower in skip_values or val_lower == "skip" or val_lower == "n/a":
                return None

            # Try to convert to int
            try:
                num_val = int(val) if isinstance(
                    val, (int, float)) else int(float(val))
                if num_val in skip_values:
                    return None
                return num_val
            except (ValueError, TypeError):
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
                    # Skip if this duct has a stop value
                    if has_stop_value(connected_duct):
                        output.print_md("*Stopping at {} (ID: {}) - has 'stop' value*".format(
                            connected_duct.family if connected_duct.family else "Unknown",
                            output.linkify(connected_duct.element.Id)
                        ))
                        continue
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
    """Check if duct has a skip value in its number parameter, or is a round boot tap."""
    # Check if this is a round boot tap - skip those always
    family = duct.family
    family_lower = family.lower() if family else ""
    if family_lower == "boot tap":
        sig = _size_signature(duct.size)
        if sig is not None and sig[0] == "round":
            return True

    # Get all parameters and search case-insensitively
    for param in duct.element.Parameters:
        param_name_lower = param.Definition.Name.strip().lower()
        if param_name_lower not in number_paramters:
            continue

        val = param.AsString()
        if val is None:
            val = param.AsValueString()

        if val is not None:
            # Check as lowercase string
            val_lower = str(val).strip().lower()
            if val_lower in skip_values or val_lower == "skip":
                return True
            # Also check raw value in case it's an integer
            try:
                if int(val) in skip_values:
                    return True
            except (ValueError, TypeError):
                pass
    return False


def has_stop_value(duct):
    """Check if duct has a stop value in its number parameter."""
    # Get all parameters and search case-insensitively
    for param in duct.element.Parameters:
        param_name_lower = param.Definition.Name.strip().lower()
        if param_name_lower not in number_paramters:
            continue

        val = param.AsString()
        if val is None:
            val = param.AsValueString()

        if val is not None:
            val_lower = str(val).strip().lower()
            if val_lower in stop_values:
                return True
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


def find_connected_numbered_element(duct, doc, view):
    """
    Find a connected element that has a number assigned.
    For store_families (taps), look for elements connected to size_out (smaller size).
    Returns (number, duct) or (None, None) if not found.
    """
    # Check if this is a store_family
    family = duct.family
    family_lower = family.lower() if family else ""
    is_store = family_lower in store_families

    # Get all connected elements
    connected = get_connected_fittings(duct, doc, view)

    if is_store and duct.size_out:
        # For taps, prefer elements matching the smaller size (size_out)
        output.print_md("  *Looking for element matching size_out: {}*".format(duct.size_out))

        for conn in connected:
            # Check if connected element's size matches our size_out
            conn_size = conn.size
            if conn_size and duct.size_out:
                if sizes_match(duct.size_out, conn_size):
                    num = get_item_number(conn)
                    if num is not None and num > 0:
                        output.print_md("  *Found numbered element on size_out: {} (number: {})*".format(
                            conn.family if conn.family else "Unknown",
                            num
                        ))
                        return (num, conn)

    # Fallback or non-store elements: check all connected elements
    for conn in connected:
        num = get_item_number(conn)
        if num is not None and num > 0:
            output.print_md("  *Found numbered element: {} (number: {})*".format(
                conn.family if conn.family else "Unknown",
                num
            ))
            return (num, conn)

    return (None, None)


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


def number_branch_recursive(
    start_duct,
    start_number,
    doc,
    view,
    visited,
    all_stored_branches,
    modified_ducts,
        filter_by_size=None,
        skip_start_numbering=False):
    """
    Number a branch starting from start_duct with start_number.
    Processes depth-first: if we encounter more store_families, process those sub-branches first.
    filter_by_size: If provided, only process connected elements matching this size (for filtering branches)
    Returns the last number used.
    """
    current_number = start_number

    # Number the start duct (optional)
    if not skip_start_numbering:
        if is_numberable(start_duct) and not has_skip_value(start_duct):
            set_item_number(start_duct, current_number)
            modified_ducts.append(start_duct)
            output.print_md("  Branch: Set {} to **{}** (ID: {})".format(
                start_duct.family if start_duct.family else "Unknown",
                current_number,
                output.linkify(start_duct.element.Id)
            ))
            current_number += 1

    visited.add(start_duct.id)

    # Get connected elements and process them
    to_process = []
    connected = get_connected_fittings(start_duct, doc, view)
    apply_size_filter = True

    for conn in connected:
        if conn.id in visited:
            continue

        family = conn.family
        family_lower = family.lower() if family else ""

        # If this is a store_family, always collect as a sub-branch (size may differ)
        if family_lower in store_families:
            # Skip round boot taps - never add them to branches
            if has_skip_value(conn):
                output.print_md("    *Skipping {} - round boot tap*".format(
                    family if family else "Unknown"
                ))
            else:
                all_stored_branches.append(conn)
                output.print_md("    *Found sub-branch: {} (ID: {})*".format(
                    family if family else "Unknown",
                    output.linkify(conn.element.Id)
                ))
            continue

        # If we're filtering by size, skip non-store elements that don't match
        # Only apply on the first hop from the tap to choose branch direction
        if filter_by_size and apply_size_filter:
            conn_size = conn.size
            if conn_size:
                if not sizes_match(filter_by_size, conn_size):
                    output.print_md("    *Skipping {} - size {} doesn't match filter {}*".format(
                        conn.family if conn.family else "Unknown",
                        conn_size,
                        filter_by_size
                    ))
                    continue

        # Only process if numberable or traversable
        if is_numberable(conn) or is_traversable(conn):
            to_process.append(conn)

    # After the first hop, do not size-filter deeper traversal
    apply_size_filter = False

    # Process all connected elements (breadth-first on this level)
    for duct in to_process:
        if duct.id in visited:
            continue

        visited.add(duct.id)

        # Number if numberable
        if is_numberable(duct) and not has_skip_value(duct):
            set_item_number(duct, current_number)
            modified_ducts.append(duct)
            output.print_md("  Branch: Set {} to **{}** (ID: {})".format(
                duct.family if duct.family else "Unknown",
                current_number,
                output.linkify(duct.element.Id)
            ))
            current_number += 1

        # Continue down this path
        next_connected = get_connected_fittings(duct, doc, view)
        for next_conn in next_connected:
            if next_conn.id not in visited:
                family = next_conn.family
                family_lower = family.lower() if family else ""

                # If store_family, add as sub-branch (ignore size filter)
                if family_lower in store_families:
                    # Skip round boot taps - never add them to branches
                    if has_skip_value(next_conn):
                        output.print_md("    *Skipping {} - round boot tap*".format(
                            family if family else "Unknown"
                        ))
                    else:
                        all_stored_branches.append(next_conn)
                        output.print_md("    *Found sub-branch: {} (ID: {})*".format(
                            family if family else "Unknown",
                            output.linkify(next_conn.element.Id)
                        ))
                else:
                    if is_numberable(next_conn) or is_traversable(next_conn):
                        to_process.append(next_conn)

    return current_number - 1


def number_run_forward(
        start_duct,
        start_number,
        doc,
        view,
        visited=None,
        stored_taps=None,
        modified_ducts=None,
        allow_store_families=False,
        filter_by_size=None):
    """
    Number fittings sequentially starting from start_duct with start_number.
    Simply increments the number for each numberable fitting (no duplicate matching).
    allow_store_families: If True, store_families can be numbered (used when they are selected)
    filter_by_size: If provided, only process elements with sizes containing this string
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

    # Apply size filter if provided
    if filter_by_size:
        output.print_md("*Filtering by size: {}*".format(filter_by_size))
        filtered_connected = []
        for conn in connected:
            conn_size = conn.size_in if conn.size_in else ""
            if sizes_match(filter_by_size, conn_size):
                filtered_connected.append(conn)
            else:
                output.print_md("  *Skipped {} (size {} doesn't match filter)*".format(
                    conn.family if conn.family else "Unknown",
                    conn_size
                ))
        connected = filtered_connected

    output.print_md("*Found {} connected fittings*".format(len(connected)))

    to_process = [(conn, current_number)
                  for conn in connected if conn.id not in visited]

    while to_process:
        duct, num = to_process.pop(0)

        if duct.id in visited:
            output.print_md("  *Skipping {} - already visited in this run*".format(
                duct.family if duct.family else "Unknown"
            ))
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
            # Skip round boot taps - don't even store them
            if has_skip_value(duct):
                output.print_md("  *Skipping {} - has skip value*".format(
                    family if family else "Unknown"
                ))
                continue

            if not allow_store_families:
                stored_taps.append((duct, None))
                output.print_md("  *Stored tap for later (no value written)*")
                # Skip during traversal unless it was the selected fitting
                continue
            # If allow_store_families is True, fall through to number it

        # Check if we should number this fitting
        if is_numberable(duct):
            if has_skip_value(duct):
                output.print_md("  *Skipping {} - has skip value, not modifying*".format(
                    family if family else "Unknown"
                ))
                # Do NOT modify - just skip it completely
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
                selected_duct.family if selected_duct.family else "Unknown",
                start_number,
                output.linkify(selected_duct.element.Id)
            ))

            output.print_md("---")

            # Continue numbering forward from selected fitting
            stored_taps = []
            allow_stores = is_selected_store_family if 'is_selected_store_family' in locals() else False

            # If starting from a store_family (tap), filter by size_out
            filter_size = None
            if is_selected_store_family and selected_duct.size_out:
                filter_size = selected_duct.size_out
                output.print_md(
                    "*Starting from tap - will only number elements matching size_out: {}*".format(filter_size))

            last_number, stored_taps, forward_modified, forward_count = number_run_forward(
                selected_duct,
                start_number + 1,
                doc,
                view,
                visited,
                stored_taps,
                [],
                allow_stores,
                filter_size
            )
            modified_ducts.extend(forward_modified)
            modified_count = 1 + forward_count

            # Process stored branches recursively
            if stored_taps:
                output.print_md("---")
                output.print_md(
                    "## Processing {} stored branches...".format(len(stored_taps)))

                # Process each stored branch
                branches_to_process = [tap_duct for tap_duct, _ in stored_taps]

                while branches_to_process:
                    branch_duct = branches_to_process.pop(0)

                    # Stored taps were marked visited during main run traversal,
                    # but we still need to process their branches here.
                    if branch_duct.id in visited and not (
                        branch_duct.family and branch_duct.family.lower() in store_families
                    ):
                        continue

                    # Find connected numbered element to get anchor
                    anchor_num, anchor_duct = find_connected_numbered_element(branch_duct, doc, view)

                    if anchor_num is None:
                        output.print_md("*Warning: Could not find anchor for branch {} (ID: {})*".format(
                            branch_duct.family if branch_duct.family else "Unknown",
                            output.linkify(branch_duct.element.Id)
                        ))
                        continue

                    # Round up to nearest 10 after the last used number (one branch at a time)
                    base_for_branch = (last_number + 1) if last_number is not None else (anchor_num + 1)
                    branch_start = round_up_to_nearest_10(base_for_branch)

                    output.print_md("### Branch from {} (anchor: {}, rounded: {})".format(
                        branch_duct.family if branch_duct.family else "Unknown",
                        anchor_num,
                        branch_start
                    ))

                    # For store_families, only process elements matching size_out
                    filter_size = branch_duct.size_out if branch_duct.family and branch_duct.family.lower() in store_families else None
                    if filter_size:
                        output.print_md("  *Filtering branch by size_out: {}*".format(filter_size))

                    # Number this branch and collect any sub-branches (use fresh list for each branch)
                    sub_branches = []

                    # Check if tap should be skipped (e.g., round boot taps)
                    if not has_skip_value(branch_duct):
                        # Set the tap itself to the bucket number (e.g., 100)
                        set_item_number(branch_duct, branch_start)
                        modified_ducts.append(branch_duct)
                        output.print_md("  Branch: Set {} to **{}** (ID: {})".format(
                            branch_duct.family if branch_duct.family else "Unknown",
                            branch_start,
                            output.linkify(branch_duct.element.Id)
                        ))
                    else:
                        output.print_md("  Branch: Skipping {} - has skip value (ID: {})".format(
                            branch_duct.family if branch_duct.family else "Unknown",
                            output.linkify(branch_duct.element.Id)
                        ))

                    # Start numbering connected branch elements at bucket+1 (e.g., 101)
                    branch_first = branch_start + 1
                    branch_last = number_branch_recursive(
                        branch_duct,
                        branch_first,
                        doc,
                        view,
                        visited,
                        sub_branches,
                        modified_ducts,
                        filter_size,
                        skip_start_numbering=True
                    )

                    # Update last number if this branch went higher
                    if branch_last > last_number:
                        last_number = branch_last

                    # Add sub-branches to the front of the queue (depth-first)
                    # This ensures C's sub-branches are processed before D
                    if sub_branches:
                        output.print_md("  *Found {} sub-branches - adding to queue*".format(len(sub_branches)))
                        branches_to_process = sub_branches + branches_to_process

                    output.print_md("  *Branch complete. Last number: {}*".format(branch_last))

                modified_count = len(modified_ducts)

            output.print_md("---")
            output.print_md(
                "## Numbering complete! Last number used: **{}** | Total modified: **{}**".format(last_number, len(modified_ducts)))

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
