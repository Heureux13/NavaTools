# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

import math
import re

from config.parameters_registry *
from Autodesk.Revit.DB.import FabricationPart, StorageType
from pyrevit import revit, script

from ducts.revit_duct import RevitDuct
from geometry.size import Size

#fmt: off
#autopep8: off
revit_host  = globals().get("__revit__")
app         = revit_host.Application if revit_host else None
uidoc       = revit_host.ActiveUIDocument if revit_host else None
doc         = getattr(revit, "doc", None)
output      = script.get_output()
#fmt: on
# autopep8: on

number_value_parameters = [
    PYT_NUMBER_FABRICATION.lower(),
    RVT_ITEM_NUMBER.lower(),
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


# fmt:off
# autopep8: off
class RevitRuns(object):
    # Run numbering helpers wrapped as an instance API

    def __init__(
        self,
        doc_obj=None,
        view_obj=None,
        output_obj=None,
        number_parameters=None,
        skip_parameters=None,
        stop_parameters=None,
        numberable_families=None,
        traversable_families=None,
        skip_value_set=None,
        stop_value_set=None,
        stored_families=None,
        rectangle_only=True
    ):
        self.doc                        = doc_obj                   or getattr(revit, "doc", None)
        self.view                       = view_obj                  or getattr(revit, "active_view", None)
        self.output                     = output_obj                or output
        self.number_value_parameters    = list(number_parameters    or number_value_parameters)
        self.skip_check_parameters      = list(skip_parameters      or skip_check_parameters)
        self.stop_check_parameters      = list(stop_parameters      or stop_check_parameters)
        self.number_families            = set(numberable_families   or number_families)
        self.allow_but_not_number       = set(traversable_families  or allow_but_not_number)
        self.skip_values                = set(skip_value_set        or skip_values)
        self.stop_values                = set(stop_value_set        or stop_values)
        self.branch_start_families      = set(stored_families       or branch_start_families)
        self.rectangle_only             = rectangle_only
# fmt:on
# autopep8: on

    def round_up_to_nearest_10(self, number):
        # Round up to the nearest 10th
        return int(math.ceil(number / 10.0) * 10)

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

    def sizes_match(self, target_size, conn_size):
        # Return True if a size match, ignoring quotes and width/height order
        sig_a = self._size_signature(target_size)
        sig_b = self._size_signature(conn_size)

        if sig_a is None or sig_b is None:
            return False

        return sig_a == sig_b

    def get_prioritized_parameters(self, duct, parameter_names):
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

    def _has_control_value(self, duct, parameter_names, skip_values):
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
        if self.rectangle_only:
            family = duct.family
            family_lower = family.lower() if family else ""

            if family_lower in boot_families_to_skip:
                sig = self._size_signature(duct.size)

                if sig is not None and sig[0] == "round":
                    return True

        return self._has_control_value(
            duct,
            self.skip_check_parameters,
            self.skip_values,
        )

    def has_stop_value(self, duct):
        # Checks to see if duct has a stop value
        return self._has_control_value(
            duct,
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

    def set_item_number(self, duct, number):
        # Set the item number in the first available parameter
        for param in self.get_number_parameters(duct):
            if param.IsReadOnly:
                continue

            try:
                storage_type = param.StorageType
                st = storage_type

                if st == StorageType.String:
                    param.Set(str(number))
                    return True
                if st == StorageType.Integer:
                    param.Set(int(number))
                    return True
                if st == StorageType.Double:
                    param.Set(float(number))
                    return True

            except Exception:
                continue

        return False

    def get_connected_fittings(self, duct, doc_obj=None, view_obj=None):
        # Get all immediate directly connected fittings
        doc_obj = doc_obj or self.doc
        view_obj = view_obj or self.view

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
                        doc_obj,
                        view_obj,
                        connected_el,
                    )
                    if self.has_stop_value(connected_duct):
                        continue
                    connected.append(connected_duct)
                except Exception:
                    continue

        return connected

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

        family = duct.family if duct.family else ""
        signature.append(family.lower())

        size = duct.size if duct.size else ""
        signature.append(size)

        length = duct.length
        signature.append("" if length is None else str(length))

        angle = duct.angle
        signature.append("" if angle is None else str(angle))

        return tuple(signature)

    def find_duct_with_number(self, connected_ducts, target_number):
        # Find a connected fitting with a specific number
        for duct in connected_ducts:
            if self.get_item_number(duct) == target_number:
                return duct

        return None

    def follow_number_chain(self, start_duct, doc_obj=None, view_obj=None, visited=None):
        # Follow the existing number chain from start fitting.
        # Return (last_duct_in_chain, last_number_in_chain, visited_in_chain, chain_ducts).
        doc_obj = doc_obj or self.doc
        view_obj = view_obj or self.view

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
            connected = self.get_connected_fittings(
                current_duct,
                doc_obj,
                view_obj
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
                       doc_obj=None,
                       view_obj=None,
                       visited=None
                       ):
        # Find all fittings in the runt that are true endpoints only 1 connection totla
        # Returns a list of dut objects that are endpoints
        doc_obj = doc_obj or self.doc
        view_obj = view_obj or self.view

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

            connected = self.get_connected_fittings(duct, doc_obj, view_obj)
            for conn in connected:
                if conn.id not in visited and self.is_traversable(conn):
                    to_process.append(conn)

        for duct in all_ducts:
            connected = self.get_connected_fittings(duct, doc_obj, view_obj)
            traversable_count = sum(
                1 for conn in connected if self.is_traversable(conn)
            )
            if traversable_count == 1:
                endpoints.append(duct)

        return endpoints

    def find_connected_numbered_elements(self,
                                         duct,
                                         doc_obj=None,
                                         view_obj=None
                                         ):
        # Find a connected element that has a number assigned.
        # For branch_start_families (taps), look for elements connected to size_out(smaller size).
        # returns (number, duct) or (None, None) if not found
        doc_obj = doc_obj or self.doc
        doc_view = view_obj or self.view

        family = duct.family
        clean_family = family.lower().strip() if family else ""
        is_store = clean_family in self.branch_start_families

        connected = self.get_connected_fittings(duct, doc_obj, view_obj)

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
                           doc_obj=None,
                           view_obj=None,
                           visited=None
                           ):
        # Recursively search backwards through connections to find an exisitng number.
        # Returns (anchor_number, anchor_duct) or (None, None) if no anchor found.
        doc_obj = doc_obj or self.doc
        view_obj = view_obj or self.view

        if visited is None:
            visited = set()

        visited.add(duct.id)

        current_number = self.get_item_number(duct)
        if current_number is not None and current_number > 0:
            return (current_number,
                    duct
                    )

        connected = self.get_connected_fittings(duct,
                                                doc_obj,
                                                view_obj
                                                )
        for conn in connected:
            if conn.id in visited:
                continue
            if not self.is_traversable(conn):
                continue

            anchor_num, anchor_duct = self.find_anchor_number(conn,
                                                              doc_obj=doc_obj,
                                                              view_obj=view_obj,
                                                              visited=visited,
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

    def create_run_and_brach_sets(self,
                                  start_duct,
                                  doc_obj=None,
                                  view_obj=None,
                                  visited=None,
                                  branch_list=None,
                                  filter_by_size=None,
                                  ):
        # Classify immediate neighbors as run candidates or branch starts.

        doc_obj = doc_obj or self.doc
        view_obj = view_obj or self.view
        visited = visited if visited is not None else set()
        branch_list = branch_list if branch_list is not None else set()
        to_run = set()

        # Mark only the current duct as visited. Neighbors are only classified here.
        visited.add(start_duct.id)

        connected = self.get_connected_fittings(start_duct, doc_obj, view_obj)

        for conn in connected:
            if conn.id in visited:
                continue

            if filter_by_size and conn.size:
                if not self.sizes_match(filter_by_size, conn.size):
                    continue

            if self.is_rect_branch_fitting(conn):
                branch_list.add(conn)
                visited.add(conn.id)
                continue

            if self.is_traversable(conn):
                to_run.add(conn)
                visited.add(conn.id)

        return (to_run, branch_list)

    def number_branch_recursive(
        self,
        start_duct,
        start_number,
        doc_obj=None,
        view_obj=None,
        visited=None,
        all_stored_branches=None,
        modified_ducts=None,
        filter_by_size=None,
        skip_start_numbering=None,
    ):
        # Number a branch starting from start_duct with start_number.
        # Process depth first: If we encounter more branch_start_families, process those sub-branches first
        # filter_by_size: If provided, only process connected elements matching this size.
        # Returns the last number used.

        doc_obj = doc_obj or self.doc
        view_obj = view_obj or self.view
        visited = visited if visited is not None else set()
        all_stored_branches = all_stored_branches if all_stored_branches is not None else []
        modified_ducts = modified_ducts if modified_ducts is not None else []

        current_number = start_number

        if not skip_start_numbering:
            if self.is_numberable(start_duct) and not self.has_skip_value(start_duct):
                self.set_item_number(start_duct, current_number)
                modified_ducts.append(start_duct)
                current_number += 1

        visited.add(start_duct.id)

        to_process = []
        connected = self.get_connected_fittings(start_duct, doc_obj, view_obj)
        apply_size_filter = True

        for conn in connected:
            if conn.id in visited:
                continue

            family = conn.family
            cleaned_family = family.lower().strip() if family else ""

            if cleaned_family in self.branch_start_families:
                if self.has_skip_value(conn):
                    pass
                else:
                    all_stored_branches.append(conn)

            if filter_by_size and apply_size_filter:
                conn_size = conn.size
                if conn_size and not self.size_match(filter_by_size, conn_size):
                    continue

            elif self.is_numberable(conn) or self.is_traversable(conn):
                to_process.append(conn)

        for duct in to_process:
            if duct.id in visited:
                continue

            visited.add(duct.id)

            if self.is_numberable(duct) and not self.has_skip_value(duct):
                self.set_item_number(duct, current_number)
                modified_ducts.append(duct)
                current_number += 1

            next_connected = self.get_connected_fittings(
                duct,
                doc_obj,
                view_obj,
            )
            for next_conn in next_connected:
                if next_conn.id in visited:
                    continue

                family = next_conn.family
                cleaned_family = family.lower().strip() if family else ""

                if cleaned_family in self.branch_start_families:
                    if self.has_skip_value(next_conn):
                        pass
                    else:
                        all_stored_branches.append(next_conn)

                if filter_by_size and apply_size_filter:
                    conn_size = conn.size
                    if conn_size and not self.size_match(filter_by_size, conn_size):
                        continue

                elif self.is_numberable(next_conn) or self.is_traversable(next_conn):
                    to_process.append(next_conn)

        return current_number - 1
