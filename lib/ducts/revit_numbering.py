# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

import math
import re

from config.parameters_registry import (
    PYT_NUMBER_ORDER,
    PYT_NUMBER_FABRICATION,
    PYT_SKIP_NUMBER,
    RVT_ANGLE,
    RVT_FAMILY,
    RVT_ITEM_NUMBER,
    RVT_LENGTH,
    RVT_SIZE,
)
from Autodesk.Revit.DB import FabricationPart, StorageType
from pyrevit import revit, script

from ducts.revit_duct import RevitDuct
from geometry.size import Size

#fmt: off
#autopep8: off
revit_host  = globals().get("__revit__")
app         = revit_host.Application if revit_host else None
uidoc       = revit_host.ActiveUIDocument if revit_host else None
doc         = getattr(revit, "doc", None)
view        = getattr(revit, "active_view", None)
output      = script.get_output()
#fmt: on
# autopep8: on

number_value_parameters = [
    RVT_ITEM_NUMBER.lower(),
    PYT_NUMBER_FABRICATION.lower(),
]

skip_check_parameters = [
    PYT_SKIP_NUMBER.lower(),
]

stop_check_parameters = list(number_value_parameters)

match_parameters = {
    RVT_FAMILY.lower(),
    RVT_SIZE.lower(),
    RVT_LENGTH.lower(),
    RVT_ANGLE.lower(),
}

number_families = {
    "drop cheek",
    "elbow",
    "elbow - 90 degree",
    "end cap",
    "offset",
    "ogee",
    "radius elbow",
    "reducer",
    "square to ø",
    "straight",
    "tdf end cap",
    "tee",
    "transition",
}

allow_but_not_number = {
    "access door",
    "canvas",
    "fire damper - type a",
    "fire damper - type b",
    "fire damper - type c",
    "fire damper - type cr",
    "manbars",
    "rect volume damper",
    "smoke fire damper - type cr",
    "smoke fire damper - type csr",
    "straight tap",
}

skip_values = {
    0,
    "skip",
    "n/a",
}

stop_values = {
    "stop"
}

branch_start_families = {
    "boot tap",
    "straight tap",
    "rec on rnd straight tap"
}

boot_families_to_skip = {
    "boot tap",
}


# fmt:off
# autopep8: off
class RevitNumbers(object):
    # Run numbering helpers wrapped as an instance API

    def __init__(
        self,
        output_obj                  =None,
        number_parameters           =None,
        skip_parameters             =None,
        stop_parameters             =None,
        numberable_families         =None,
        traversable_families        =None,
        skip_value_set              =None,
        stop_value_set              =None,
        stored_families             =None,
        allow_rectangle             =True,
        allow_round                 =True,
        allow_oval                  =True,
    ):
        self.doc                        = getattr(revit, "doc", None)
        self.view                       = getattr(revit, "active_view", None)
        self.output                     = output_obj                or output
        self.number_value_parameters    = list(number_parameters    or number_value_parameters)
        self.skip_check_parameters      = list(skip_parameters      or skip_check_parameters)
        self.stop_check_parameters      = list(stop_parameters      or stop_check_parameters)
        self.number_families            = set(numberable_families   or number_families)
        self.allow_but_not_number       = set(traversable_families  or allow_but_not_number)
        self.skip_values                = set(skip_value_set        or skip_values)
        self.stop_values                = set(stop_value_set        or stop_values)
        self.branch_start_families      = set(stored_families       or branch_start_families)
        self.allow_rectangle            = allow_rectangle
        self.allow_round                = allow_round
        self.allow_oval                 = allow_oval
# fmt:on
# autopep8: on

    def round_up_to_nearest_10(self, number):
        # Round up to the nearest 10
        return int(math.ceil(number / 10.0) * 10)

    def round_up_to_nearest_100(self, number):
        # Round up to the nearest 100
        return int(math.ceil(number / 100.0) * 100)

    def round_up_to_nearest_1000(self, number):
        # Round up to the nearest 1000
        return int(math.ceil(number / 1000.0) * 1000)

    def _size_signature(self, size_value):
        # Returns normalied shpae and size
        if size_value is None:
            return None

        size_str = str(size_value).strip()
        if not size_str:
            return None

        size_obj = Size(size_str)

        if size_obj.in_diameter is not None:
            return ("round", round(float(size_obj.in_diameter), 4))

        if size_obj.in_oval_dia is not None:
            width = size_obj.in_width
            height = size_obj.in_height

            if width is not None and height is not None:
                return ("oval", round(float(width), 4), round(float(height), 4))

        if size_obj.in_width is not None and size_obj.in_height is not None:
            dims = sorted(
                [round(float(size_obj.in_width), 4),
                 round(float(size_obj.in_height), 4)]
            )
            return ("rect", tuple(dims))

        return None

    def is_rectangular_size(self, size_value):
        # Checks if a size is a rectangle
        sig = self._size_signature(size_value)

        return sig is not None and sig[0] == "rect"

    def is_oval_size(self, size_value):
        # Checks if a size is an oval
        sig = self._size_signature(size_value)

        return sig is not None and sig[0] == "oval"

    def is_round_size(self, size_value):
        # Check if a size is round
        sig = self._size_signature(size_value)

        return sig is not None and sig[0] == "round"

    def is_allowed_shape(self, duct):
        sig = self._size_signature(duct.size)

        if sig is None:
            return None
        if sig[0] == "rect":
            return self.allow_rectangle
        if sig[0] == "round":
            return self.allow_round
        if sig[0] == "oval":
            return self.allow_oval

        return True

    def sizes_match(self,
                    target_size,
                    conn_size
                    ):
        # Return True if a size match, ignoring quotes and width/height order
        sig_a = self._size_signature(target_size)
        sig_b = self._size_signature(conn_size)

        if sig_a is None or sig_b is None:
            return False

        return sig_a == sig_b

    def get_prioritized_parameters(self,
                                   duct,
                                   parameter_names
                                   ):
        # Return matching parameters in the configured priority order
        cleaned = [n.strip().lower() for n in parameter_names]
        dic = {n: [] for n in cleaned}

        for d in duct.element.Parameters:
            pname = d.Definition.Name.strip().lower()
            if pname in dic:
                dic[pname].append(d)

        ordered_params = []
        for name in cleaned:
            ordered_params.extend(dic.get(name, []))

        return ordered_params

    def get_number_parameters(self, duct):
        # Return item number parameters in configured read/write priority order
        return self.get_prioritized_parameters(duct, self.number_value_parameters)

    @staticmethod
    def _get_parameter_value(param):
        # Return a parameter value as a string when possible
        value = param.AsString()

        if value is None:
            value = param.AsValueString()

        return value

    def _has_control_value(self,
                           duct,
                           parameter_names,
                           skip_values
                           ):
        # Return True when any configured control parametres contain a control value
        for param in self.get_prioritized_parameters(duct, parameter_names):
            value = self._get_parameter_value(param)

            if value is None:
                continue

            value_lower = str(value).strip().lower()
            if value_lower in skip_values:
                return True

        return False

    def has_skip_value(self, duct):
        # Check if duct has a skip value in its number parameter or is a round boot taop
        if not self.allow_round:
            family = duct.family
            family_lower = family.lower() if family else ""

            if family_lower in boot_families_to_skip:
                sig = self._size_signature(duct.size)

                if sig is not None and sig[0] == "round":
                    return True

        return self._has_control_value(duct,
                                       self.skip_check_parameters,
                                       self.skip_values,
                                       )

    def has_stop_value(self, duct):
        # Checks to see if duct has a stop value
        return self._has_control_value(duct,
                                       self.stop_check_parameters,
                                       self.stop_values,
                                       )

    def get_item_number(self, duct):
        # Get the current item number form any of the number parameters
        if self.has_skip_value(duct):
            return None

        for param in self.get_number_parameters(duct):
            value = self._get_parameter_value(param)

            if value is None:
                continue

            try:
                return int(value) if isinstance(value, (int, float)) else int(float(value))
            except (ValueError, TypeError):
                match = re.search(r'\d+', str(value))
                if match:
                    return int(match.group())

        return None

    def get_order_numbers(self, view_obj=None, scope="view"):
        # Return RevitDuct items with a value in PYT_NUMBER_ORDER.
        # NOTE: PYT_NUMBER_ORDER is read-only sequencing input. It is never written.
        # scope="view" limits to active/provided view; scope="project" scans all views.
        scope_value = str(scope).strip().lower(
        ) if scope is not None else "view"
        use_project_scope = scope_value in ("project", "all", "all_project")
        target_view = None if use_project_scope else (view_obj or self.view)

        ducts = RevitDuct.all(self.doc, target_view)

        filtered = []
        for duct in ducts:
            try:
                param = duct.element.LookupParameter(PYT_NUMBER_ORDER)
                if not param:
                    continue

                # HasValue is reliable for non-string parameter types.
                try:
                    if not param.HasValue:
                        continue
                except Exception:
                    pass

                value = param.AsString()
                if value is None:
                    value = param.AsValueString()

                if value is not None and str(value).strip() != "":
                    filtered.append(duct)
            except Exception:
                continue

        return filtered

    def get_order_number_text(self, duct):
        # Return PYT_NUMBER_ORDER as trimmed text, or empty string when missing
        if not duct:
            return ""

        try:
            param = duct.element.LookupParameter(PYT_NUMBER_ORDER)
            if not param:
                return ""

            value = param.AsString()
            if value is None:
                value = param.AsValueString()

            return "" if value is None else str(value).strip()
        except Exception:
            return ""

    def find_duplicate_order_numbers(self, ducts):
        # Return map of order values that appear more than once
        grouped = {}

        for duct in ducts or []:
            value = self.get_order_number_text(duct)
            if not value:
                continue

            if value not in grouped:
                grouped[value] = []
            grouped[value].append(duct)

        return {k: v for k, v in grouped.items() if len(v) > 1}

    def has_any_item_number(self, ducts):
        # Return True when at least one duct already has an item number
        for duct in ducts or []:
            if self.get_item_number(duct) is not None:
                return True

        return False

    def number_ordered_runs(self,
                            ordered_ducts,
                            first_start_number=None,
                            repeat_numbers=False,
                            ):
        # Number runs by PYT_NUMBER_ORDER from low to high.
        # NOTE: PYT_NUMBER_ORDER is only used to determine run order.
        #       This method modifies item numbers only.
        # Start at the lowest order number (or override with first_start_number).
        # After each run: next_start = round_up_to_nearest_100(last_used + 100).
        if not ordered_ducts:
            return []

        ordered_pairs = []
        for duct in ordered_ducts:
            if not duct:
                continue

            try:
                param = duct.element.LookupParameter(PYT_NUMBER_ORDER)
                if not param:
                    continue

                value = param.AsString()
                if value is None:
                    value = param.AsValueString()

                if value is None:
                    continue

                text = str(value).strip()
                if not text:
                    continue

                try:
                    order_value = int(text)
                except Exception:
                    try:
                        num = float(text)
                        int_num = int(num)
                        if num != int_num:
                            continue
                        order_value = int_num
                    except Exception:
                        continue

                ordered_pairs.append((order_value, duct))
            except Exception:
                continue

        if not ordered_pairs:
            return []

        ordered_pairs.sort(key=lambda x: x[0])
        sorted_ducts = [duct for _, duct in ordered_pairs]

        if first_start_number is None:
            first_ordered_duct = ordered_pairs[0][1]
            existing_item = self.get_item_number(first_ordered_duct)
            next_start_number = existing_item if existing_item is not None else ordered_pairs[
                0][0]
        else:
            next_start_number = int(first_start_number)

        results = []
        processed_ids = set()

        for idx, start_duct in enumerate(sorted_ducts):
            if start_duct.id in processed_ids:
                continue

            try:
                last_used_number, run_piece_count, visited_ids = self._number_run_simple(
                    start_duct,
                    next_start_number,
                    repeat_numbers=repeat_numbers,
                )
                processed_ids.update(visited_ids)
                results.append(
                    (start_duct, next_start_number, last_used_number))

                if run_piece_count > 500:
                    next_start_number = self.round_up_to_nearest_1000(
                        last_used_number + 1000)
                else:
                    next_start_number = self.round_up_to_nearest_100(
                        last_used_number + 100)
            except Exception:
                pass

        return results

    def _number_run_simple(self, start_duct, start_number, repeat_numbers=False):
        """Map-based run numbering with stored-branch queue behavior."""
        connectivity_map = self.build_connectivity_map(start_duct)

        visited = set([start_duct.id])
        modified_ducts = []
        stored_branches = []
        current_number = start_number
        piece_count = 0
        previous_signature = None

        start_family = start_duct.family.lower() if start_duct.family else ""
        start_is_branch = start_family in self.branch_start_families

        if (self.is_numberable(start_duct) or start_is_branch) and not self.has_skip_value(start_duct):
            assigned_number, current_number, previous_signature = self.assign_number_by_signature(
                start_duct,
                current_number,
                previous_signature,
                repeat_numbers=repeat_numbers,
            )
            self.set_item_number(start_duct, assigned_number)
            modified_ducts.append(start_duct)
            piece_count += 1

        last_used_number = self.get_item_number(start_duct)
        if last_used_number is None:
            last_used_number = current_number - 1

        filter_size = start_duct.size_out if start_is_branch and start_duct.size_out else None
        last_used_number, stored_branches, forward_modified, _ = self._number_run_forward_map(
            start_duct,
            current_number,
            visited=visited,
            stored_taps=stored_branches,
            modified_ducts=[],
            allow_store_families=start_is_branch,
            filter_by_size=filter_size,
            connectivity_map=connectivity_map,
            repeat_numbers=repeat_numbers,
            previous_signature=previous_signature,
        )
        modified_ducts.extend(forward_modified)
        piece_count += len(forward_modified)

        branches_to_process = list(stored_branches)
        while branches_to_process:
            branch_duct, stored_anchor_duct = branches_to_process.pop(0)

            branch_family = branch_duct.family.lower() if branch_duct.family else ""
            if branch_duct.id in visited and branch_family not in self.branch_start_families:
                continue

            anchor_num = None
            anchor_duct = stored_anchor_duct
            if anchor_duct is not None:
                anchor_num = self.get_item_number(anchor_duct)

            if anchor_num is None:
                anchor_num, anchor_duct = self.find_connected_numbered_element(
                    branch_duct,
                    connectivity_map=connectivity_map,
                )

            if anchor_num is None:
                continue

            base_for_branch = (
                last_used_number + 1) if last_used_number is not None else (anchor_num + 1)
            branch_start = self.round_up_to_nearest_10(base_for_branch)

            branch_filter_size = branch_duct.size_out if branch_family in self.branch_start_families else None
            sub_branches = []

            if not self.has_skip_value(branch_duct):
                self.set_item_number(branch_duct, branch_start)
                modified_ducts.append(branch_duct)
                piece_count += 1

            branch_first = branch_start + 1
            branch_last, branch_piece_count = self._number_branch_recursive_map(
                branch_duct,
                branch_first,
                visited,
                sub_branches,
                modified_ducts,
                filter_by_size=branch_filter_size,
                skip_start_numbering=True,
                connectivity_map=connectivity_map,
                repeat_numbers=repeat_numbers,
            )
            piece_count += branch_piece_count

            if branch_last > last_used_number:
                last_used_number = branch_last

            if sub_branches:
                branches_to_process = sub_branches + branches_to_process

        return last_used_number, piece_count, visited

    def _number_run_forward_map(self,
                                start_duct,
                                start_number,
                                visited=None,
                                stored_taps=None,
                                modified_ducts=None,
                                allow_store_families=False,
                                filter_by_size=None,
                                connectivity_map=None,
                                repeat_numbers=False,
                                previous_signature=None,
                                ):
        if visited is None:
            visited = set()
        if stored_taps is None:
            stored_taps = []
        if modified_ducts is None:
            modified_ducts = []

        current_number = start_number
        current_signature = previous_signature
        last_assigned_number = start_number - 1

        connected = self.get_connected_from_map(start_duct,
                                                connectivity_map=connectivity_map,
                                                )
        if filter_by_size:
            filtered_connected = []
            for conn in connected:
                conn_size = conn.size_in if conn.size_in else ""
                if self.sizes_match(filter_by_size, conn_size):
                    filtered_connected.append(conn)
            connected = filtered_connected

        to_process = [(conn, start_duct)
                      for conn in connected if conn.id not in visited]
        max_iterations = 10000
        iterations = 0

        while to_process and iterations < max_iterations:
            iterations += 1
            duct, source_duct = to_process.pop(0)

            if duct.id in visited:
                continue

            visited.add(duct.id)

            family = duct.family
            family_lower = family.lower() if family else ""

            if family_lower in self.branch_start_families:
                if self.has_skip_value(duct):
                    continue

                if not allow_store_families:
                    stored_taps.append((duct, source_duct))
                    continue

            if self.is_numberable(duct) or family_lower in self.branch_start_families:
                if not self.has_skip_value(duct):
                    assigned_number, current_number, current_signature = self.assign_number_by_signature(
                        duct,
                        current_number,
                        current_signature,
                        repeat_numbers=repeat_numbers,
                    )
                    self.set_item_number(duct, assigned_number)
                    last_assigned_number = assigned_number
                    modified_ducts.append(duct)
            elif not self.is_traversable(duct):
                continue

            next_connected = self.get_connected_from_map(duct,
                                                         connectivity_map=connectivity_map,
                                                         )
            for conn in next_connected:
                if conn.id not in visited:
                    to_process.append((conn, duct))

        return last_assigned_number, stored_taps, modified_ducts, len(modified_ducts)

    def _number_branch_recursive_map(self,
                                     start_duct,
                                     start_number,
                                     visited,
                                     all_stored_branches,
                                     modified_ducts,
                                     filter_by_size=None,
                                     skip_start_numbering=False,
                                     connectivity_map=None,
                                     repeat_numbers=False,
                                     ):
        current_number = start_number
        piece_count = 0
        previous_signature = None
        last_assigned_number = start_number - 1

        if not skip_start_numbering:
            start_family = start_duct.family.lower() if start_duct.family else ""
            if (self.is_numberable(start_duct) or start_family in self.branch_start_families) and not self.has_skip_value(start_duct):
                assigned_number, current_number, previous_signature = self.assign_number_by_signature(
                    start_duct,
                    current_number,
                    previous_signature,
                    repeat_numbers=repeat_numbers,
                )
                self.set_item_number(start_duct, assigned_number)
                last_assigned_number = assigned_number
                modified_ducts.append(start_duct)
                piece_count += 1

        visited.add(start_duct.id)

        to_process = []
        connected = self.get_connected_from_map(start_duct,
                                                connectivity_map=connectivity_map,
                                                )
        apply_size_filter = True

        for conn in connected:
            if conn.id in visited:
                continue

            family = conn.family
            family_lower = family.lower() if family else ""

            if family_lower in self.branch_start_families:
                if not self.has_skip_value(conn):
                    all_stored_branches.append((conn, start_duct))
                continue

            if filter_by_size and apply_size_filter:
                conn_size = conn.size
                if conn_size and not self.sizes_match(filter_by_size, conn_size):
                    continue

            if self.is_numberable(conn) or self.is_traversable(conn):
                to_process.append(conn)

        apply_size_filter = False
        max_iterations = 10000
        iterations = 0

        while to_process and iterations < max_iterations:
            iterations += 1
            duct = to_process.pop(0)
            if duct.id in visited:
                continue

            visited.add(duct.id)

            if self.is_numberable(duct) and not self.has_skip_value(duct):
                assigned_number, current_number, previous_signature = self.assign_number_by_signature(
                    duct,
                    current_number,
                    previous_signature,
                    repeat_numbers=repeat_numbers,
                )
                self.set_item_number(duct, assigned_number)
                last_assigned_number = assigned_number
                modified_ducts.append(duct)
                piece_count += 1

            next_connected = self.get_connected_from_map(duct,
                                                         connectivity_map=connectivity_map,
                                                         )
            for next_conn in next_connected:
                if next_conn.id in visited:
                    continue

                family = next_conn.family
                family_lower = family.lower() if family else ""

                if family_lower in self.branch_start_families:
                    if not self.has_skip_value(next_conn):
                        all_stored_branches.append((next_conn, duct))
                elif self.is_numberable(next_conn) or self.is_traversable(next_conn):
                    to_process.append(next_conn)

        return last_assigned_number, piece_count

    def _find_connected_numbered_element(self, duct):
        """Find a connected element that has a number assigned."""
        family = duct.family
        family_lower = family.lower() if family else ""
        is_branch = family_lower in self.branch_start_families

        connected = self._get_connected_fittings(duct)

        # For branches, prefer elements matching size_out
        if is_branch and hasattr(duct, 'size_out') and duct.size_out:
            for conn in connected:
                conn_size = conn.size if hasattr(conn, 'size') else None
                if conn_size and self.sizes_match(duct.size_out, conn_size):
                    num = self.get_item_number(conn)
                    if num is not None and num > 0:
                        return (num, conn)

        # Fallback: check all connected elements
        for conn in connected:
            num = self.get_item_number(conn)
            if num is not None and num > 0:
                return (num, conn)

        return (None, None)

    def set_item_number(self, duct, number):
        # Set only item-number parameters. Never write PYT_NUMBER_ORDER.
        updated = False

        for param in self.get_number_parameters(duct):
            if param.IsReadOnly:
                continue

            param_name = param.Definition.Name if param and param.Definition else ""
            if param_name and param_name.strip().lower() == PYT_NUMBER_ORDER.lower():
                continue

            try:
                storage_type = param.StorageType
                st = storage_type

                if st == StorageType.String:
                    param.Set(str(number))
                    updated = True
                    continue
                if st == StorageType.Integer:
                    param.Set(int(number))
                    updated = True
                    continue
                if st == StorageType.Double:
                    param.Set(float(number))
                    updated = True
                    continue

            except Exception:
                continue

        return updated

    def _get_connected_fittings(self, duct):
        # Query Revit connectors directly for immediate neighbors
        connected = []
        for connector in duct.get_connectors():
            if not connector.IsConnected:
                continue

            for ref in list(connector.AllRefs):
                if not ref or not hasattr(ref, "Owner"):
                    continue

                connected_el = ref.Owner
                if not isinstance(connected_el, FabricationPart):
                    continue
                if connected_el.Id == duct.element.Id:
                    continue

                try:
                    connected_duct = RevitDuct(
                        self.doc,
                        self.view,
                        connected_el,
                    )
                    if self.has_stop_value(connected_duct):
                        continue
                    connected.append(connected_duct)
                except Exception:
                    continue

        return connected

    def build_connectivity_map(self, start_duct):
        # Build a full adjacency map once so downstream traversal avoids repeated API scans
        connectivity_map = {}
        to_process = [start_duct]
        visited = set()

        while to_process:
            current = to_process.pop(0)
            if current.id in visited:
                continue

            visited.add(current.id)

            neighbors = self._get_connected_fittings(current)
            connectivity_map[current.id] = neighbors

            for neighbor in neighbors:
                if neighbor.id not in visited:
                    to_process.append(neighbor)

        return connectivity_map

    def get_connected_from_map(self,
                               duct,
                               connectivity_map=None,
                               ):
        # Read neighbors from an optional prebuilt map and fallback to direct lookup
        if connectivity_map is None:
            return self._get_connected_fittings(duct)

        return connectivity_map.get(duct.id, [])

    def is_numberable(self, duct):
        # Check if a duct can be numbered based on family
        family = duct.family
        if not family:
            return False
        return family.lower() in self.number_families

    def is_traversable(self, duct):
        # Checks if we can traverse through the duct
        family = duct.family
        if not family:
            return False
        family_lower = family.lower()
        return family_lower in self.allow_but_not_number or self.is_numberable(duct)

    def get_match_signature(self, duct):
        # Get the match signature for a duct based on match paramters
        # Returns a tuple of (family, size, length, angle) for comparison

        signature = []

        family = (duct.family or "").strip().lower()
        signature.append(family)

        size = duct.size if duct.size else ""
        size_signature = self._size_signature(size)
        signature.append(
            size_signature if size_signature is not None else str(size).strip())

        length = duct.length
        if length is None:
            signature.append("")
        else:
            # Snap to 1/16" to avoid floating noise while still matching by length.
            length_value = float(length)
            signature.append(round(length_value * 16.0) / 16.0)

        angle = duct.angle
        if angle is None:
            signature.append("")
        else:
            signature.append(round(float(angle), 3))

        return tuple(signature)

    def get_repeat_match_signature(self, duct):
        # Duplicate mode uses the full signature, including length.
        return self.get_match_signature(duct)

    def find_duct_with_number(self,
                              connected_ducts,
                              target_number
                              ):
        # Find a connected fitting with a specific number
        for duct in connected_ducts:
            if self.get_item_number(duct) == target_number:
                return duct

        return None

    def follow_number_chain(self,
                            start_duct,
                            visited=None,
                            connectivity_map=None,
                            ):
        # Follow the existing number chain from start fitting.
        # Return (last_duct_in_chain, last_number_in_chain, visited_in_chain, chain_ducts).
        if visited is None:
            visited = set()

        chain_ducts = []
        current_duct = start_duct
        current_number = self.get_item_number(current_duct)

        if current_number is None:
            return (current_duct, None, visited, chain_ducts)

        visited.add(current_duct.id)
        chain_ducts.append(current_duct)

        while True:
            next_number = current_number + 1
            connected = self.get_connected_from_map(current_duct,
                                                    connectivity_map=connectivity_map,
                                                    )
            unvisited_traversable = [
                conn for conn in connected
                if conn.id not in visited and self.is_traversable(conn)
            ]

            next_duct = self.find_duct_with_number(
                unvisited_traversable,
                next_number
            )
            if next_duct is None:
                break

            visited.add(next_duct.id)
            chain_ducts.append(next_duct)

            current_duct = next_duct
            current_number = next_number

        return (current_duct, current_number, visited, chain_ducts)

    def find_endpoints(self,
                       start_duct,
                       visited=None,
                       connectivity_map=None,
                       ):
        # Find all fittings in the runt that are true endpoints only 1 connection totla
        # Returns a list of dut objects that are endpoints
        if visited is None:
            visited = set()

        endpoints = []
        all_ducts = []
        to_process = [start_duct]

        while to_process:
            duct = to_process.pop(0)
            if duct.id in visited:
                continue

            visited.add(duct.id)
            all_ducts.append(duct)

            connected = self.get_connected_from_map(duct,
                                                    connectivity_map=connectivity_map,
                                                    )
            for conn in connected:
                if conn.id not in visited and self.is_traversable(conn):
                    to_process.append(conn)

        for duct in all_ducts:
            connected = self.get_connected_from_map(duct,
                                                    connectivity_map=connectivity_map,
                                                    )
            traversable_count = sum(
                1 for conn in connected if self.is_traversable(conn)
            )
            if traversable_count == 1:
                endpoints.append(duct)

        return endpoints

    def find_connected_numbered_element(self,
                                        duct,
                                        connectivity_map=None,
                                        ):
        # Find a connected element that has a number assigned.
        # For branch_start_families (taps), look for elements connected to size_out(smaller size).
        # returns (number, duct) or (None, None) if not found
        family = duct.family
        clean_family = family.lower().strip() if family else ""
        is_store = clean_family in self.branch_start_families

        connected = self.get_connected_from_map(duct,
                                                connectivity_map=connectivity_map,
                                                )

        if is_store and duct.size_out:
            for conn in connected:
                conn_size = conn.size
                if conn_size and duct.size_out and self.sizes_match(duct.size_out, conn_size):
                    number = self.get_item_number(conn)
                    if number is not None and number > 0:
                        return (number, conn)

        for conn in connected:
            number = self.get_item_number(conn)
            if number is not None and number > 0:
                return (number, conn)

        return (None, None)

    def find_anchor_number(self,
                           duct,
                           visited=None,
                           connectivity_map=None,
                           ):
        # Recursively search backwards through connections to find an exisitng number.
        # Returns (anchor_number, anchor_duct) or (None, None) if no anchor found.
        if visited is None:
            visited = set()

        visited.add(duct.id)

        current_number = self.get_item_number(duct)
        if current_number is not None and current_number > 0:
            return (current_number,
                    duct
                    )

        connected = self.get_connected_from_map(duct,
                                                connectivity_map=connectivity_map,
                                                )
        for conn in connected:
            if conn.id in visited:
                continue
            if not self.is_traversable(conn):
                continue

            anchor_num, anchor_duct = self.find_anchor_number(conn,
                                                              visited=visited,
                                                              connectivity_map=connectivity_map,
                                                              )
            if anchor_num is not None:
                return (anchor_num, anchor_duct)

        return (None, None)

    def is_rect_branch_fitting(self,
                               fitting
                               ):
        # Returns true if a fitting is a start duct fitting
        family = fitting.family
        cleaned_family = family.lower().strip() if family else ""

        if cleaned_family not in self.branch_start_families:
            return False

        size_in = fitting.size_in
        size_out = fitting.size_out

        if not size_in or not size_out:
            return False

        return (
            self.is_rectangular_size(size_in) and
            self.is_rectangular_size(size_out)
        )

    def collect_run_and_branch_sets(self,
                                    start_duct,
                                    visited=None,
                                    branch_list=None,
                                    filter_by_size=None,
                                    connectivity_map=None,
                                    ):
        # Classify immediate neighbors as run candidates or branch starts.
        # Uses connectivity map when provided to avoid repeated API scans.
        visited = visited if visited is not None else set()
        branch_list = branch_list if branch_list is not None else []

        to_run = []
        to_run_ids = set()
        branch_ids = set([b.id for b in branch_list])
        to_process = [start_duct]
        max_iterations = 10000

        iteration = 0
        while to_process and iteration < max_iterations:
            iteration += 1
            current = to_process.pop(0)

            if current.id in visited:
                continue
            visited.add(current.id)

            if current.id not in to_run_ids:
                to_run.append(current)
                to_run_ids.add(current.id)

            connected = self.get_connected_from_map(current,
                                                    connectivity_map=connectivity_map,
                                                    )

            for conn in connected:
                if conn.id in visited:
                    continue

                if filter_by_size and conn.size:
                    if not self.sizes_match(filter_by_size, conn.size):
                        continue

                if self.is_rect_branch_fitting(conn) and conn.id not in branch_ids:
                    branch_ids.add(conn.id)
                    branch_list.append(conn)

                elif self.is_traversable(conn) and conn.id not in to_run_ids:
                    to_run.append(conn)
                    to_run_ids.add(conn.id)
                    to_process.append(conn)

        return (to_run, branch_list)

    def assign_number_by_signature(self,
                                   duct,
                                   current_number,
                                   previous_signature,
                                   repeat_numbers=False):

        current_signature = self.get_match_signature(duct)

        if not repeat_numbers:
            assigned_number = current_number
            current_number += 1
        else:
            current_signature = self.get_repeat_match_signature(duct)

            if previous_signature is None or current_signature == previous_signature:
                assigned_number = current_number
            else:
                current_number += 1
                assigned_number = current_number

        return assigned_number, current_number, current_signature
