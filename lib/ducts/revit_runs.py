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


revit_host = globals().get("__revit__")
app = revit_host.Application if revit_host else None
uidoc = revit_host.ActiveUIDocument if revit_host else None
doc = getattr(revit, "doc", None)
view = getattr(revit, "active_view", None)
output = script.get_output()


number_value_parameters = [
    PYT_NUMBER_FABRICATION.lower(),
    RVT_ITEM_NUMBER.lower(),
]

skip_check_parameters = [
    PYT_SKIP_NUMBER.lower(),
]

stop_check_parameters = list(number_value_parameters)

match_paramters = {
    RVT_FAMILY.lower(),
    RVT_SIZE.lower(),
    RVT_LENGTH.lower(),
    RVT_ANGLE.lower(),
}

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
    "reducer",
    "tee",
}

allow_but_not_number = {
    "manbars",
    "canvas",
    "fire damper - type a",
    "fire damper - type b",
    "fire damper - type c",
    "fire damper - type cr",
    "smoke fire damper - type cr",
    "smoke fire damper - type csr",
    "rect volume damper",
    "access door",
    "straight tap",
}

skip_values = {
    0,
    "skip",
    "n/a",
}

stop_values = {
    "stop",
}

branch_start_families = {
    "boot tap",
    "straight tap",
    "rec on rnd straight tap",
}


class RevitRuns(object):
    """Run-numbering helpers wrapped as an instance API."""

    def __init__(
        self,
        doc_obj=None,
        view_obj=None,
        output_obj=None,
        number_parameters=None,
        skip_parameters=None,
        stop_parameters=None,
        numberable_families=None,
        traversable_familie=None,
        skip_value_set=None,
        stop_value_set=None,
        stored_families=None,
    ):
        # fmt: off
        # autopep8: off
        self.doc                            = doc_obj or getattr(revit, "doc", None)
        self.view                           = view_obj or getattr(revit, "active_view", None)
        self.output                         = output_obj or output
        self.number_value_parameters        = list(number_parameters or number_value_parameters)
        self.skip_check_parameters          = list(skip_parameters or skip_check_parameters)
        self.stop_check_parameters          = list(stop_parameters or stop_check_parameters)
        self.number_families                = set(numberable_families or number_families)
        self.allow_but_not_number           = set(traversable_families or allow_but_not_number)
        self.skip_values                    = set(skip_value_set or skip_values)
        self.stop_values                    = set(stop_value_set or stop_values)
        self.branch_start_families                 = set(stored_families or branch_start_families)
        # fmt: on
        # autopep8: on

    def round_up_to_nearest_10(self, number):
        """Round up to the nearest 10. E.g., 55 -> 60, 60 -> 60, 1 -> 10"""
        return int(math.ceil(number / 10.0) * 10)

    def _size_signature(self, size_value):
        """Create a normalized size signature for matching sizes."""
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
        """Check if a size is rectangular (not round or oval)."""
        sig = self._size_signature(size_value)
        return sig is not None and sig[0] == "rect"

    def sizes_match(self, filter_size, conn_size):
        """Return True if sizes match, ignoring quotes and width/height order."""
        sig_a = self._size_signature(filter_size)
        sig_b = self._size_signature(conn_size)
        if sig_a is None or sig_b is None:
            return False
        return sig_a == sig_b

    def get_prioritized_parameters(self, duct, parameter_names):
        """Return matching parameters in the configured priority order."""
        params_by_name = {name: [] for name in parameter_names}

        for param in duct.element.Parameters:
            param_name_lower = param.Definition.Name.strip().lower()
            if param_name_lower in params_by_name:
                params_by_name[param_name_lower].append(param)

        ordered_params = []
        for name in parameter_names:
            ordered_params.extend(params_by_name.get(name, []))

        return ordered_params

    def get_number_parameters(self, duct):
        """Return item number parameters in configured read/write priority order."""
        return self.get_prioritized_parameters(duct, self.number_value_parameters)

    @staticmethod
    def _get_parameter_value(param):
        """Return a parameter value as a string when possible."""
        value = param.AsString()
        if value is None:
            value = param.AsValueString()
        return value

    def _has_control_value(self, duct, parameter_names, control_values):
        """Return True when any configured control parameter contains a control value."""
        for param in self.get_prioritized_parameters(duct, parameter_names):
            value = self._get_parameter_value(param)
            if value is None:
                continue

            value_lower = str(value).strip().lower()
            if value_lower in control_values:
                return True

            try:
                if int(value) in control_values:
                    return True
            except (ValueError, TypeError):
                pass

        return False
#

    def get_item_number(self, duct):
        """Get the current item number from any of the number parameters."""
        if self.has_skip_value(duct):
            return None

        for param in self.get_number_parameters(duct):
            value = self._get_parameter_value(param)
            if value is None:
                continue

            try:
                return int(value) if isinstance(value, (int, float)) else int(float(value))
            except (ValueError, TypeError):
                match = re.search(r"\d+", str(value))
                if match:
                    return int(match.group())

        return None

    def set_item_number(self, duct, number):
        """Set the item number in the first available parameter."""
        for param in self.get_number_parameters(duct):
            if param.IsReadOnly:
                continue

            try:
                storage_type = param.StorageType
                if storage_type == StorageType.String:
                    param.Set(str(number))
                    return True
                if storage_type == StorageType.Integer:
                    param.Set(int(number))
                    return True
                if storage_type == StorageType.Double:
                    param.Set(float(number))
                    return True
            except Exception:
                continue

        return False

    def get_connected_fittings(self, duct, doc_obj=None, view_obj=None):
        """Get all immediately connected fittings (only direct connections)."""
        doc_obj = doc_obj or self.doc
        view_obj = view_obj or self.view

        connected = []
        for connector in duct.get_connectors():
            if not connector.IsConnected:
                continue

            for ref in list(connector.AllRefs):
                if not ref or not hasattr(ref, "Owner"):
                    continue

                connected_elem = ref.Owner
                if not isinstance(connected_elem, FabricationPart):
                    continue
                if connected_elem.Id == duct.element.Id:
                    continue

                try:
                    connected_duct = RevitDuct(
                        doc_obj, view_obj, connected_elem)
                    if self.has_stop_value(connected_duct):
                        continue
                    connected.append(connected_duct)
                except Exception:
                    continue

        return connected

    def is_numberable(self, duct):
        """Check if a duct can be numbered based on family."""
        family = duct.family
        if not family:
            return False
        return family.lower() in self.number_families

    def is_traversable(self, duct):
        """Check if we can traverse through this duct (even if not numbering it)."""
        family = duct.family
        if not family:
            return False
        family_lower = family.lower()
        return family_lower in self.allow_but_not_number or self.is_numberable(duct)

    def has_skip_value(self, duct):
        """Check if duct has a skip value in its number parameter, or is a round boot tap."""
        family = duct.family
        family_lower = family.lower() if family else ""
        if family_lower == "boot tap":
            sig = self._size_signature(duct.size)
            if sig is not None and sig[0] == "round":
                return True

        return self._has_control_value(
            duct,
            self.skip_check_parameters,
            self.skip_values,
        )

    def has_stop_value(self, duct):
        """Check if duct has a stop value in its number parameter."""
        return self._has_control_value(
            duct,
            self.stop_check_parameters,
            self.stop_values,
        )

    def get_match_signature(self, duct):
        """
        Get the match signature for a duct based on match_parameters.
        Returns a tuple of (family, size, length, angle) for comparison.
        """
        signature = []

        family = duct.family if duct.family else ""
        signature.append(family.lower())

        size = duct.size if hasattr(duct, "size") and duct.size else ""
        signature.append(str(size))

        length = ""
        try:
            for param in duct.element.Parameters:
                param_name_lower = param.Definition.Name.strip().lower()
                if param_name_lower == "length":
                    value = self._get_parameter_value(param)
                    if value:
                        length = str(value)
                    break
        except Exception:
            pass
        signature.append(length)

        angle = ""
        try:
            for param in duct.element.Parameters:
                param_name_lower = param.Definition.Name.strip().lower()
                if param_name_lower == "angle":
                    value = self._get_parameter_value(param)
                    if value:
                        angle = str(value)
                    break
        except Exception:
            pass
        signature.append(angle)

        return tuple(signature)

    def find_duct_with_number(self, connected_ducts, target_number):
        """Find a connected fitting with a specific number."""
        for duct in connected_ducts:
            if self.get_item_number(duct) == target_number:
                return duct
        return None

    def follow_number_chain(self, start_duct, doc_obj=None, view_obj=None, visited=None):
        """
        Follow the existing number chain from the start fitting.
        Returns (last_duct_in_chain, last_number_in_chain, visited_in_chain, chain_ducts).
        """
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
                current_duct, doc_obj, view_obj)
            unvisited_traversable = [
                conn for conn in connected
                if conn.id not in visited and self.is_traversable(conn)
            ]

            next_duct = self.find_duct_with_number(
                unvisited_traversable, next_number)
            if next_duct is None:
                break

            visited.add(next_duct.id)
            chain_ducts.append(next_duct)
            current_duct = next_duct
            current_number = next_number

        return (current_duct, current_number, visited, chain_ducts)

    def find_endpoints(self, start_duct, doc_obj=None, view_obj=None, visited=None):
        """
        Find all fittings in the run that are true endpoints (only 1 traversable connection total).
        Returns a list of duct objects that are endpoints.
        """
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
                1 for conn in connected if self.is_traversable(conn))
            if traversable_count == 1:
                endpoints.append(duct)

        return endpoints

    def find_connected_numbered_element(self, duct, doc_obj=None, view_obj=None):
        """
        Find a connected element that has a number assigned.
        For branch_start_families (taps), look for elements connected to size_out (smaller size).
        Returns (number, duct) or (None, None) if not found.
        """
        doc_obj = doc_obj or self.doc
        view_obj = view_obj or self.view

        family = duct.family
        family_lower = family.lower() if family else ""
        is_store = family_lower in self.branch_start_families

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

    def find_anchor_number(self, duct, doc_obj=None, view_obj=None, visited=None):
        """
        Recursively search backwards through connections to find an existing number.
        Returns (anchor_number, anchor_duct) or (None, None) if no anchor found.
        """
        doc_obj = doc_obj or self.doc
        view_obj = view_obj or self.view
        if visited is None:
            visited = set()

        visited.add(duct.id)

        current_number = self.get_item_number(duct)
        if current_number is not None and current_number > 0:
            return (current_number, duct)

        connected = self.get_connected_fittings(duct, doc_obj, view_obj)
        for conn in connected:
            if conn.id in visited:
                continue
            if not self.is_traversable(conn):
                continue

            anchor_num, anchor_duct = self.find_anchor_number(
                conn,
                doc_obj=doc_obj,
                view_obj=view_obj,
                visited=visited,
            )
            if anchor_num is not None:
                return (anchor_num, anchor_duct)

        return (None, None)

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
        skip_start_numbering=False,
    ):
        """
        Number a branch starting from start_duct with start_number.
        Processes depth-first: if we encounter more branch_start_families, process those sub-branches first.
        filter_by_size: If provided, only process connected elements matching this size.
        Returns the last number used.
        """
        # fmt: off
        # autopep8: off
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
            family_lower = family.lower() if family else ""

            if family_lower in self.branch_start_families:
                if self.has_skip_value(conn):
                    pass
                else:
                    all_stored_branches.append(conn)
                continue

            if filter_by_size and apply_size_filter:
                conn_size = conn.size
                if conn_size and not self.sizes_match(filter_by_size, conn_size):
                    continue

            if self.is_numberable(conn) or self.is_traversable(conn):
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
                duct, doc_obj, view_obj)
            for next_conn in next_connected:
                if next_conn.id in visited:
                    continue

                family = next_conn.family
                family_lower = family.lower() if family else ""

                if family_lower in self.branch_start_families:
                    if self.has_skip_value(next_conn):
                        pass
                    else:
                        all_stored_branches.append(next_conn)
                elif self.is_numberable(next_conn) or self.is_traversable(next_conn):
                    to_process.append(next_conn)

        return current_number - 1

    def number_run_forward(
        self,
        start_duct,
        start_number,
        doc_obj=None,
        view_obj=None,
        visited=None,
        stored_taps=None,
        modified_ducts=None,
        allow_branch_start_families=False,
        filter_by_size=None,
    ):
        """
        Number fittings sequentially starting from start_duct with start_number.
        Simply increments the number for each numberable fitting (no duplicate matching).
        allow_branch_start_families: If True, branch_start_families can be numbered.
        filter_by_size: If provided, only process elements with matching sizes.
        Returns the last number used, list of stored tap fittings, and modified ducts.
        """
        doc_obj = doc_obj or self.doc
        view_obj = view_obj or self.view
        visited = visited if visited is not None else set()
        stored_taps = stored_taps if stored_taps is not None else []
        modified_ducts = modified_ducts if modified_ducts is not None else []

        current_number = start_number
        connected = self.get_connected_fittings(start_duct, doc_obj, view_obj)

        if filter_by_size:
            filtered_connected = []
            for conn in connected:
                conn_size = conn.size_in if conn.size_in else ""
                if self.sizes_match(filter_by_size, conn_size):
                    filtered_connected.append(conn)
            connected = filtered_connected

        to_process = [(conn, current_number)
                      for conn in connected if conn.id not in visited]

        while to_process:
            duct, _ = to_process.pop(0)

            if duct.id in visited:
                continue

            visited.add(duct.id)

            family = duct.family
            family_lower = family.lower() if family else ""

            if family_lower in self.branch_start_families:
                if self.has_skip_value(duct):
                    continue
                if not allow_branch_start_families:
                    stored_taps.append((duct, None))
                    continue

            if self.is_numberable(duct):
                if self.has_skip_value(duct):
                    pass
                else:
                    self.set_item_number(duct, current_number)
                    modified_ducts.append(duct)
                    current_number += 1
            elif not self.is_traversable(duct):
                continue

            next_connected = self.get_connected_fittings(
                duct, doc_obj, view_obj)
            for conn in next_connected:
                if conn.id not in visited:
                    to_process.append((conn, current_number))

        return current_number - 1, stored_taps, modified_ducts, len(modified_ducts)


_default_runs = RevitRuns()


def round_up_to_nearest_10(number):
    return _default_runs.round_up_to_nearest_10(number)


def _size_signature(size_value):
    return _default_runs._size_signature(size_value)


def is_rectangular_size(size_value):
    return _default_runs.is_rectangular_size(size_value)


def sizes_match(filter_size, conn_size):
    return _default_runs.sizes_match(filter_size, conn_size)


def get_prioritized_parameters(duct, parameter_names):
    return _default_runs.get_prioritized_parameters(duct, parameter_names)


def get_number_parameters(duct):
    return _default_runs.get_number_parameters(duct)


def _get_parameter_value(param):
    return RevitRuns._get_parameter_value(param)


def _has_control_value(duct, parameter_names, control_values):
    return _default_runs._has_control_value(duct, parameter_names, control_values)


def get_item_number(duct):
    return _default_runs.get_item_number(duct)


def set_item_number(duct, number):
    return _default_runs.set_item_number(duct, number)


def get_connected_fittings(duct, doc, view):
    return _default_runs.get_connected_fittings(duct, doc, view)


def is_numberable(duct):
    return _default_runs.is_numberable(duct)


def is_traversable(duct):
    return _default_runs.is_traversable(duct)


def has_skip_value(duct):
    return _default_runs.has_skip_value(duct)


def has_stop_value(duct):
    return _default_runs.has_stop_value(duct)


def get_match_signature(duct):
    return _default_runs.get_match_signature(duct)


def find_duct_with_number(connected_ducts, target_number):
    return _default_runs.find_duct_with_number(connected_ducts, target_number)


def follow_number_chain(start_duct, doc, view, visited=None):
    return _default_runs.follow_number_chain(start_duct, doc, view, visited)


def find_endpoints(start_duct, doc, view, visited=None):
    return _default_runs.find_endpoints(start_duct, doc, view, visited)


def find_connected_numbered_element(duct, doc, view):
    return _default_runs.find_connected_numbered_element(duct, doc, view)


def find_anchor_number(duct, doc, view, visited=None):
    return _default_runs.find_anchor_number(duct, doc, view, visited)


def number_branch_recursive(
    start_duct,
    start_number,
    doc,
    view,
    visited,
    all_stored_branches,
    modified_ducts,
    filter_by_size=None,
    skip_start_numbering=False,
):
    return _default_runs.number_branch_recursive(
        start_duct,
        start_number,
        doc,
        view,
        visited,
        all_stored_branches,
        modified_ducts,
        filter_by_size,
        skip_start_numbering,
    )


def number_run_forward(
    start_duct,
    start_number,
    doc,
    view,
    visited=None,
    stored_taps=None,
    modified_ducts=None,
    allow_branch_start_families=False,
    filter_by_size=None,
):
    return _default_runs.number_run_forward(
        start_duct,
        start_number,
        doc,
        view,
        visited,
        stored_taps,
        modified_ducts,
        allow_branch_start_families,
        filter_by_size,
    )
