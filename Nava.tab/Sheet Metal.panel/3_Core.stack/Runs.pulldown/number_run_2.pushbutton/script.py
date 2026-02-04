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
__title__ = "Numbers All w/ Matches"
__doc__ = """
Will number all fittings in job with matching numbers if possible
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
# 1. Scan entire model for all FabricationPart elements
# 2. Group numberable ducts into connected runs (BFS traversal)
# 3. Number each run sequentially, respecting signatures for deduplication
# 4. Allow numbers to spread across runs if identical signatures exist
# 5. Process stored_families afterwards, matching connected signatures where possible

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
    # Strip whitespace and asterisks from family name
    family_clean = family.strip().strip('*').lower()
    return family_clean in number_families


def is_traversable(duct):
    """Check if we can traverse through this duct (even if not numbering it)."""
    family = duct.family
    if not family:
        return False
    # Strip whitespace and asterisks from family name
    family_clean = family.strip().strip('*').lower()
    return (family_clean in allow_but_not_number or
            is_numberable(duct))


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


def collect_connected_run(start_duct, doc, view, visited_global=None):
    """
    Collect all ducts in a connected run starting from start_duct.
    Uses BFS to find all connected traversable ducts.
    Returns list of all ducts in this run.
    """
    if visited_global is None:
        visited_global = set()

    run_ducts = []
    to_process = [start_duct]

    while to_process:
        duct = to_process.pop(0)

        if duct.id in visited_global:
            continue
        visited_global.add(duct.id)
        run_ducts.append(duct)

        # Get all connected fittings
        connected = get_connected_fittings(duct, doc, view)
        for conn in connected:
            if conn.id not in visited_global and is_traversable(conn):
                to_process.append(conn)

    return run_ducts, visited_global


def get_all_numberable_and_traversable(doc, view):
    """
    Collect all numberable and traversable ducts from the entire model.
    Also collect stored families separately for processing at the end.
    Returns (numberable_ducts, all_traversable_ducts, stored_family_ducts).
    """
    collector = FilteredElementCollector(doc, view.Id)
    all_fab_parts = collector.OfClass(FabricationPart).ToElements()

    numberable_ducts = []
    traversable_ducts = []
    stored_family_ducts = []

    for elem in all_fab_parts:
        try:
            duct = RevitDuct(doc, view, elem)
            family = duct.family
            family_clean = family.strip().strip('*').lower() if family else ""

            if family_clean in store_families:
                stored_family_ducts.append(duct)
            elif is_numberable(duct):
                numberable_ducts.append(duct)
            if is_traversable(duct):
                traversable_ducts.append(duct)
        except Exception:
            continue

    return numberable_ducts, traversable_ducts, stored_family_ducts


def organize_into_runs(all_traversable_ducts, doc, view):
    """
    Organize all traversable ducts into connected runs.
    Returns a list of runs, where each run is a list of ducts.
    """
    visited_global = set()
    runs = []

    for duct in all_traversable_ducts:
        if duct.id not in visited_global:
            run_ducts, visited_global = collect_connected_run(
                duct, doc, view, visited_global)
            runs.append(run_ducts)

    return runs


def get_duplicate_signatures(all_ducts):
    """
    Find all signatures that appear more than once.
    Returns a dict: {signature: [list of ducts with that signature]}
    """
    signature_groups = {}

    for duct in all_ducts:
        if not is_numberable(duct) or has_skip_value(duct):
            continue

        sig = get_match_signature(duct)
        if sig not in signature_groups:
            signature_groups[sig] = []
        signature_groups[sig].append(duct)

    # Filter to only duplicates (signatures with 2+ ducts)
    duplicates = {sig: ducts for sig,
                  ducts in signature_groups.items() if len(ducts) > 1}
    return duplicates


def number_duplicates(all_numberable_ducts, signature_map, current_number):
    """
    Find all duplicate signatures and pre-number them by size (largest to smallest).
    Returns (next_available_number, modified_ducts_list, updated_signature_map).
    """
    duplicates = get_duplicate_signatures(all_numberable_ducts)
    modified_ducts = []
    next_number = current_number

    if not duplicates:
        output.print_md("*No duplicates found*")
        return next_number, modified_ducts, signature_map

    output.print_md(
        "## Found {} duplicate signatures, numbering by size (largest first)".format(len(duplicates)))

    for sig, duct_list in duplicates.items():
        # Sort by size - largest first
        # Size is the second element in signature tuple
        duct_list_sorted = sorted(duct_list, key=lambda d: str(
            d.size if hasattr(d, 'size') and d.size else ""), reverse=True)

        output.print_md(
            "**Signature**: {} | Count: {}".format(sig, len(duct_list_sorted)))

        for idx, duct in enumerate(duct_list_sorted):
            set_item_number(duct, next_number)
            modified_ducts.append(duct)
            signature_map[sig] = next_number
            output.print_md("  Set {} ({}) to **{}**".format(
                duct.family if duct.family else "Unknown",
                duct.size if hasattr(
                    duct, 'size') and duct.size else "Unknown",
                next_number
            ))
            next_number += 1

        output.print_md("---")

    return next_number, modified_ducts, signature_map


def number_run_all_at_once(run_ducts, current_number, doc, view, signature_map=None, modified_ducts=None):
    """
    Number all ducts in a single run, starting from current_number.
    Uses signature_map for deduplication across runs.
    Returns (last_number_used, modified_count, updated_signature_map, updated_modified_ducts).
    """
    if signature_map is None:
        signature_map = {}
    if modified_ducts is None:
        modified_ducts = []

    visited = set()
    to_process = run_ducts[:]  # Start with all ducts in run
    max_number_used = current_number - 1
    modified_in_run = 0

    # Use BFS to number all ducts in run while keeping numbers grouped
    while to_process:
        duct = to_process.pop(0)

        if duct.id in visited:
            continue
        visited.add(duct.id)

        # Skip if not numberable
        if not is_numberable(duct):
            continue

        # Skip if has skip value
        if has_skip_value(duct):
            set_item_number(duct, "skip")
            modified_ducts.append(duct)
            modified_in_run += 1
            output.print_md("  *{} - Skip value*".format(
                duct.family if duct.family else "Unknown"
            ))
            continue

        # Check signature for matching
        duct_signature = get_match_signature(duct)

        if duct_signature in signature_map:
            # Match found - reuse number
            matching_number = signature_map[duct_signature]
            set_item_number(duct, matching_number)
            modified_ducts.append(duct)
            modified_in_run += 1
            output.print_md("Set {} to **{}** (ID: {}) *[matched]*".format(
                duct.family if duct.family else "Unknown",
                matching_number,
                output.linkify(duct.element.Id)
            ))
        else:
            # New unique element - assign next number
            set_item_number(duct, current_number)
            modified_ducts.append(duct)
            modified_in_run += 1
            signature_map[duct_signature] = current_number
            output.print_md("Set {} to **{}** (ID: {})".format(
                duct.family if duct.family else "Unknown",
                current_number,
                output.linkify(duct.element.Id)
            ))
            max_number_used = current_number
            current_number += 1

    return max_number_used, modified_in_run, signature_map, modified_ducts


# Main Script
# ==================================================

try:
    # Collect all ducts from the model
    output.print_md("## Collecting all ducts from model...")
    all_numberable, all_traversable, all_stored_families = get_all_numberable_and_traversable(
        doc, view)
    output.print_md(
        "Found **{}** numberable ducts".format(len(all_numberable)))
    output.print_md(
        "Found **{}** traversable ducts".format(len(all_traversable)))
    output.print_md(
        "Found **{}** stored family ducts".format(len(all_stored_families)))
    output.print_md("---")

    if not all_numberable:
        output.print_md("## No numberable ducts found in model")
    else:
        # Start transaction
        t = Transaction(doc, "Number All Ducts")
        t.Start()

        try:
            output.print_md("## Organizing into connected runs...")
            runs = organize_into_runs(all_traversable, doc, view)
            output.print_md("Found **{}** connected runs".format(len(runs)))
            output.print_md("---")

            # First: Number all duplicates by size (largest to smallest)
            output.print_md("## Step 1: Pre-numbering duplicates by size")
            signature_map = {}
            all_modified_ducts = []
            current_number = 1

            # Number duplicates first
            dup_next_number, dup_modified, signature_map = number_duplicates(
                all_numberable, signature_map, current_number)
            all_modified_ducts.extend(dup_modified)
            current_number = dup_next_number

            output.print_md("---")
            output.print_md("## Step 2: Numbering runs")

            # Now process all runs - duplicates already have numbers assigned
            for run_idx, run_ducts in enumerate(runs):
                output.print_md("## Run {} ({} ducts)".format(
                    run_idx + 1, len(run_ducts)))

                # Filter to only numberable in this run
                numberable_in_run = [
                    d for d in run_ducts if is_numberable(d)]

                if numberable_in_run:
                    last_num, modified_count, signature_map, all_modified_ducts = number_run_all_at_once(
                        numberable_in_run,
                        current_number,
                        doc,
                        view,
                        signature_map,
                        all_modified_ducts
                    )
                    current_number = last_num + 1
                    output.print_md("*Modified: {} ducts, Next number: {}*".format(
                        modified_count, current_number))
                else:
                    output.print_md("*No numberable ducts in this run*")

                output.print_md("---")

            # Now handle stored_families
            output.print_md(
                "## Processing stored families (boot taps, etc.)...")
            output.print_md("Found **{}** stored family ducts".format(
                len(all_stored_families)))

            stored_modified = 0
            for stored_duct in all_stored_families:
                # Try to find connected numbered duct to get signature
                connected = get_connected_fittings(stored_duct, doc, view)
                numbered_connected = [c for c in connected if get_item_number(
                    c) is not None]

                if numbered_connected:
                    # Get signature from connected numbered duct and use its number
                    ref_duct = numbered_connected[0]
                    ref_number = get_item_number(ref_duct)
                    set_item_number(stored_duct, ref_number)
                    all_modified_ducts.append(stored_duct)
                    stored_modified += 1
                    output.print_md("Set {} to **{}** (ID: {}) *[from connected run]*".format(
                        stored_duct.family if stored_duct.family else "Unknown",
                        ref_number,
                        output.linkify(stored_duct.element.Id)
                    ))
                else:
                    # No connected numbered duct - assign next sequential
                    set_item_number(stored_duct, current_number)
                    all_modified_ducts.append(stored_duct)
                    stored_modified += 1
                    current_number += 1
                    output.print_md("Set {} to **{}** (ID: {}) *[sequential]*".format(
                        stored_duct.family if stored_duct.family else "Unknown",
                        get_item_number(stored_duct),
                        output.linkify(stored_duct.element.Id)
                    ))

            if all_stored_families:
                output.print_md("---")
                output.print_md("*Processed {} stored families*".format(
                    stored_modified))

            output.print_md("---")
            output.print_md(
                "## Complete! Total modified: **{}** | Final number: **{}**".format(len(all_modified_ducts), current_number - 1))

            # Select all modified ducts
            if all_modified_ducts:
                RevitElement.select_many(uidoc, all_modified_ducts)
                output.print_md(
                    "## Selected all {} modified fittings".format(len(all_modified_ducts)))

            # Commit transaction
            t.Commit()

        except Exception as e:
            t.RollBack()
            output.print_md("## Error during numbering: {}".format(str(e)))
            import traceback
            output.print_md("```\n{}\n```".format(traceback.format_exc()))

except Exception as e:
    output.print_md("## Error: {}".format(str(e)))
    import traceback
    output.print_md("```\n{}\n```".format(traceback.format_exc()))

# Final print statements
print_disclaimer(output)
