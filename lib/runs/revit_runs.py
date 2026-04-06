# -*- coding: utf-8 -*-
"""=========================================================================
Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder.
========================================================================="""

# Standard library
# =========================================================
from ducts.revit_xyz import RevitXYZ
from ducts.revit_duct import RevitDuct
from geometry.size import Size
from geometry.offsets import Offsets
from Autodesk.Revit.DB import (
    ElementId,
    FilteredElementCollector,
    BuiltInCategory,
    UnitUtils,
    FabricationPart,
    UnitTypeId,
    ConnectorType
)
import re
import logging
import math
from enum import Enum

# Thrid Party
from pyrevit import DB, revit, script

#
import clr
clr.AddReference("RevitAPI")

# Variables
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
view = revit.active_view
output = script.get_output()


class RevitRuns(object):
    """Run utilities wrapped as instance helpers."""

    def __init__(
        self,
        doc=None,
        view=None,
        output_obj=None,
        number_paramters=None,
        skip_values=None,
        stop_values=None,
        number_families=None,
        allow_but_not_number=None,
        store_families=None,
    ):
        self.doc = doc or revit.doc
        self.view = view or revit.active_view
        self.output = output_obj or output
        self.number_paramters = set(number_paramters or [])
        self.skip_values = set(skip_values or [])
        self.stop_values = set(stop_values or [])
        self.number_families = set(number_families or [])
        self.allow_but_not_number = set(allow_but_not_number or [])
        self.store_families = set(store_families or [])

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
            dims = sorted([round(float(size_obj.in_width), 4),
                          round(float(size_obj.in_height), 4)])
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

    def get_item_number(self, duct):
        """Get the current item number from any of the number parameters."""
        for param in duct.element.Parameters:
            param_name_lower = param.Definition.Name.strip().lower()
            if param_name_lower not in self.number_paramters:
                continue

            val = param.AsString()
            if val is None:
                val = param.AsValueString()

            if val is not None:
                val_lower = str(val).strip().lower()
                if val_lower in self.skip_values or val_lower == "skip" or val_lower == "n/a":
                    return None

                try:
                    num_val = int(val) if isinstance(
                        val, (int, float)) else int(float(val))
                    if num_val in self.skip_values:
                        return None
                    return num_val
                except (ValueError, TypeError):
                    match = re.search(r"\d+", str(val))
                    if match:
                        return int(match.group())
        return None

    def set_item_number(self, duct, number):
        """Set the item number in the first available parameter."""
        for param in duct.element.Parameters:
            param_name_lower = param.Definition.Name.strip().lower()

            if param_name_lower not in self.number_paramters:
                continue

            if param.IsReadOnly:
                continue

            try:
                storage_type = param.StorageType

                if storage_type == StorageType.String:
                    param.Set(str(number))
                    return True
                elif storage_type == StorageType.Integer:
                    param.Set(int(number))
                    return True
                elif storage_type == StorageType.Double:
                    param.Set(float(number))
                    return True
            except Exception:
                continue

        return False

    def get_connected_fittings(self, duct, doc=None, view=None):
        """Get all immediately connected fittings (only direct connections)."""
        doc = doc or self.doc
        view = view or self.view
        connected = []
        for connector in duct.get_connectors():
            if not connector.IsConnected:
                continue
            all_refs = list(connector.AllRefs)
            for ref in all_refs:
                if ref and hasattr(ref, 'Owner'):
                    connected_elem = ref.Owner
                    if not isinstance(connected_elem, FabricationPart):
                        continue
                    if connected_elem.Id == duct.element.Id:
                        continue
                    try:
                        connected_duct = RevitDuct(doc, view, connected_elem)
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

        for param in duct.element.Parameters:
            param_name_lower = param.Definition.Name.strip().lower()
            if param_name_lower not in self.number_paramters:
                continue

            val = param.AsString()
            if val is None:
                val = param.AsValueString()

            if val is not None:
                val_lower = str(val).strip().lower()
                if val_lower in self.skip_values or val_lower == "skip":
                    return True
                try:
                    if int(val) in self.skip_values:
                        return True
                except (ValueError, TypeError):
                    pass
        return False

    def has_stop_value(self, duct):
        """Check if duct has a stop value in its number parameter."""
        for param in duct.element.Parameters:
            param_name_lower = param.Definition.Name.strip().lower()
            if param_name_lower not in self.number_paramters:
                continue

            val = param.AsString()
            if val is None:
                val = param.AsValueString()

            if val is not None:
                val_lower = str(val).strip().lower()
                if val_lower in self.stop_values:
                    return True
        return False

    def get_match_signature(self, duct):
        """
        Get the match signature for a duct based on match_parameters.
        Returns a tuple of (family, size, length, angle) for comparison.
        """
        signature = []

        family = duct.family if duct.family else ""
        signature.append(family.lower())

        size = duct.size if hasattr(duct, 'size') and duct.size else ""
        signature.append(str(size))

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

    def find_duct_with_number(self, connected_ducts, target_number):
        """
        Find a connected fitting with a specific number.
        Returns the duct with that number or None if not found.
        """
        for duct in connected_ducts:
            num = self.get_item_number(duct)
            if num == target_number:
                return duct
        return None

    def follow_number_chain(self, start_duct, visited=None, doc=None, view=None):
        """
        Follow the existing number chain from the start fitting.
        Returns (last_duct_in_chain, last_number_in_chain, visited_in_chain, chain_ducts).
        """
        doc = doc or self.doc
        view = view or self.view
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
            connected = self.get_connected_fittings(current_duct, doc, view)

            unvisited_traversable = [
                c for c in connected
                if c.id not in visited and self.is_traversable(c)
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

    def find_endpoints(self, start_duct, visited=None, doc=None, view=None):
        """
        Find all fittings in the run that are true endpoints (only 1 traversable connection total).
        Returns a list of duct objects that are endpoints.
        """
        doc = doc or self.doc
        view = view or self.view
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

            connected = self.get_connected_fittings(duct, doc, view)
            for conn in connected:
                if conn.id not in visited and self.is_traversable(conn):
                    to_process.append(conn)

        for duct in all_ducts:
            connected = self.get_connected_fittings(duct, doc, view)
            traversable_count = sum(
                1 for c in connected if self.is_traversable(c))

            if traversable_count == 1:
                endpoints.append(duct)

        return endpoints

    def find_connected_numbered_element(self, duct, doc=None, view=None):
        """
        Find a connected element that has a number assigned.
        For store_families (taps), look for elements connected to size_out (smaller size).
        Returns (number, duct) or (None, None) if not found.
        """
        doc = doc or self.doc
        view = view or self.view

        family = duct.family
        family_lower = family.lower() if family else ""
        is_store = family_lower in self.store_families

        connected = self.get_connected_fittings(duct, doc, view)

        if is_store and duct.size_out:
            for conn in connected:
                conn_size = conn.size
                if conn_size and duct.size_out:
                    if self.sizes_match(duct.size_out, conn_size):
                        num = self.get_item_number(conn)
                        if num is not None and num > 0:
                            return (num, conn)

        for conn in connected:
            num = self.get_item_number(conn)
            if num is not None and num > 0:
                return (num, conn)

        return (None, None)

    def find_anchor_number(self, duct, visited=None, doc=None, view=None):
        """
        Recursively search backwards through connections to find an existing number.
        Returns (anchor_number, anchor_duct) or (None, None) if no anchor found.
        """
        doc = doc or self.doc
        view = view or self.view
        if visited is None:
            visited = set()

        visited.add(duct.id)

        current_number = self.get_item_number(duct)
        if current_number is not None and current_number > 0:
            return (current_number, duct)

        connected = self.get_connected_fittings(duct, doc, view)
        for conn in connected:
            if conn.id in visited:
                continue

            if not self.is_traversable(conn):
                continue

            anchor_num, anchor_duct = self.find_anchor_number(
                conn, visited, doc, view)
            if anchor_num is not None:
                return (anchor_num, anchor_duct)

        return (None, None)

    def number_branch_recursive(
        self,
        start_duct,
        start_number,
        visited,
        all_stored_branches,
        modified_ducts,
        filter_by_size=None,
        skip_start_numbering=False,
        previous_numbers=None,
        doc=None,
        view=None,
    ):
        """
        Number a branch starting from start_duct with start_number.
        Processes depth-first: if we encounter more store_families, process those sub-branches first.
        filter_by_size: If provided, only process connected elements matching this size (for filtering branches)
        Returns the last number used.
        """
        doc = doc or self.doc
        view = view or self.view
        current_number = start_number

        if not skip_start_numbering:
            if self.is_numberable(start_duct) and not self.has_skip_value(start_duct):
                if previous_numbers is not None:
                    previous_numbers[start_duct.id] = self.get_item_number(
                        start_duct)
                if self.set_item_number(start_duct, current_number):
                    modified_ducts.append(start_duct)
                    current_number += 1

        visited.add(start_duct.id)

        to_process = []
        connected = self.get_connected_fittings(start_duct, doc, view)
        apply_size_filter = True

        for conn in connected:
            if conn.id in visited:
                continue

            family = conn.family
            family_lower = family.lower() if family else ""

            if family_lower in self.store_families:
                if self.has_skip_value(conn):
                    pass
                else:
                    all_stored_branches.append(conn)
                continue

            if filter_by_size and apply_size_filter:
                conn_size = conn.size
                if conn_size:
                    if not self.sizes_match(filter_by_size, conn_size):
                        continue

            if self.is_numberable(conn) or self.is_traversable(conn):
                to_process.append(conn)

        apply_size_filter = False

        for duct in to_process:
            if duct.id in visited:
                continue

            visited.add(duct.id)

            if self.is_numberable(duct) and not self.has_skip_value(duct):
                if previous_numbers is not None:
                    previous_numbers[duct.id] = self.get_item_number(duct)
                if self.set_item_number(duct, current_number):
                    modified_ducts.append(duct)
                    current_number += 1

            next_connected = self.get_connected_fittings(duct, doc, view)
            for next_conn in next_connected:
                if next_conn.id not in visited:
                    family = next_conn.family
                    family_lower = family.lower() if family else ""

                    if family_lower in self.store_families:
                        if self.has_skip_value(next_conn):
                            pass
                        else:
                            all_stored_branches.append(next_conn)
                    else:
                        if self.is_numberable(next_conn) or self.is_traversable(next_conn):
                            to_process.append(next_conn)

        return current_number - 1

    def number_run_forward(
        self,
        start_duct,
        start_number,
        visited=None,
        stored_taps=None,
        modified_ducts=None,
        allow_store_families=False,
        filter_by_size=None,
        previous_numbers=None,
        doc=None,
        view=None,
    ):
        """
        Number fittings sequentially starting from start_duct with start_number.
        Simply increments the number for each numberable fitting (no duplicate matching).
        allow_store_families: If True, store_families can be numbered (used when they are selected)
        filter_by_size: If provided, only process elements with sizes containing this string
        Returns the last number used, list of stored tap fittings, and modified ducts.
        """
        doc = doc or self.doc
        view = view or self.view
        if visited is None:
            visited = set()
        if stored_taps is None:
            stored_taps = []
        if modified_ducts is None:
            modified_ducts = []

        current_number = start_number

        connected = self.get_connected_fittings(start_duct, doc, view)

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
            duct, num = to_process.pop(0)

            if duct.id in visited:
                continue

            visited.add(duct.id)

            family = duct.family
            family_lower = family.lower() if family else ""

            if family_lower in self.store_families:
                if self.has_skip_value(duct):
                    continue

                if not allow_store_families:
                    stored_taps.append((duct, None))
                    continue

            if self.is_numberable(duct):
                if self.has_skip_value(duct):
                    pass
                else:
                    if previous_numbers is not None:
                        previous_numbers[duct.id] = self.get_item_number(duct)
                    if self.set_item_number(duct, current_number):
                        modified_ducts.append(duct)
                        current_number += 1
            elif self.is_traversable(duct):
                pass
            else:
                continue

            next_connected = self.get_connected_fittings(duct, doc, view)
            for conn in next_connected:
                if conn.id not in visited:
                    to_process.append((conn, current_number))

        return current_number - 1, stored_taps, modified_ducts, len(modified_ducts)

    @staticmethod
    def create_duct_run(start_duct, doc, view):
        """Find all connected ducts/fittings that match both shape and size of the starting duct."""
        run = set()
        to_visit = [start_duct]
        visited = set()

        # Preload all fabrication duct parts in the view for fallback proximity checks
        try:
            all_ducts_index = {d.id: d for d in RevitDuct.all(doc, view)}
        except Exception:
            all_ducts_index = {}

        # Parse starting duct shape and size using Size.in_shape()
        start_size_obj = Size(str(start_duct.size))

        def shape_key_from_size(size_obj):
            """Create a comparable key from Size using inlet fields only."""
            shape = size_obj.in_shape()
            if shape == "round" and size_obj.in_diameter is not None:
                return ("round", round(size_obj.in_diameter, 2))
            if shape == "oval" and size_obj.in_oval_dia is not None and size_obj.in_oval_flat is not None:
                return ("oval", round(size_obj.in_oval_dia, 2), round(size_obj.in_oval_flat, 2))
            if shape == "rectangle" and size_obj.in_width is not None and size_obj.in_height is not None:
                return ("rect", round(size_obj.in_width, 2), round(size_obj.in_height, 2))
            return ("unknown", str(size_obj.in_size))

        def shape_equals(a, b, tol=0.01):
            """Compare two shape keys with tolerance for numeric parts."""
            if not isinstance(a, tuple) or not isinstance(b, tuple):
                return a == b
            if a[0] != b[0]:
                return False
            kind = a[0]
            try:
                if kind == "round":
                    return abs(float(a[1]) - float(b[1])) <= tol
                if kind == "oval":
                    return (abs(float(a[1]) - float(b[1])) <= tol and
                            abs(float(a[2]) - float(b[2])) <= tol)
                if kind == "rect":
                    # Handle orientation-insensitive comparison for rectangle
                    aw, ah = float(a[1]), float(a[2])
                    bw, bh = float(b[1]), float(b[2])
                    direct = (abs(aw - bw) <= tol and abs(ah - bh) <= tol)
                    swapped = (abs(aw - bh) <= tol and abs(ah - bw) <= tol)
                    return direct or swapped
                if kind == "unknown":
                    return a[1] == b[1]
            except Exception:
                return a == b
            return False

        start_shape = shape_key_from_size(start_size_obj)

        def connectors_close(duct_a, duct_b, tol=1e-4):
            """Fallback: check if any connectors from two ducts are coincident within tolerance (feet)."""
            try:
                conns_a = duct_a.get_connectors() or []
                conns_b = duct_b.get_connectors() or []
            except Exception:
                return False
            for ca in conns_a:
                oa = None
                try:
                    oa = ca.Origin
                except Exception:
                    pass
                if oa is None:
                    continue
                for cb in conns_b:
                    ob = None
                    try:
                        ob = cb.Origin
                    except Exception:
                        pass
                    if ob is None:
                        continue
                    try:
                        dx = oa.X - ob.X
                        dy = oa.Y - ob.Y
                        dz = oa.Z - ob.Z
                        if (dx * dx + dy * dy + dz * dz) <= (tol * tol):
                            return True
                    except Exception:
                        continue
            return False

        while to_visit:
            duct = to_visit.pop()
            if duct.id in visited:
                continue
            visited.add(duct.id)
            run.add(duct)
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
                        try:
                            connected_duct = RevitDuct(
                                doc, view, connected_elem)
                        except Exception:
                            continue
                        # Parse connected duct shape and size via Size.in_shape()
                        connected_size_obj = Size(str(connected_duct.size))
                        connected_shape = shape_key_from_size(
                            connected_size_obj)
                        # Match by normalized shape/size only (avoid string formatting mismatches)
                        if shape_equals(connected_shape, start_shape) and connected_duct.id not in visited:
                            to_visit.append(connected_duct)
                # Fallback: if no owner references provided by API, try proximity to other parts
                if all_ducts_index:
                    for other_id, other_duct in all_ducts_index.items():
                        if other_id == duct.id or other_id in visited:
                            continue
                        # Pre-filter by shape/size to limit work
                        try:
                            other_shape = shape_key_from_size(
                                Size(str(other_duct.size)))
                        except Exception:
                            continue
                        if not shape_equals(other_shape, start_shape):
                            continue
                        if connectors_close(duct, other_duct):
                            to_visit.append(other_duct)
        return list(run)

    @staticmethod
    def create_duct_run_same_height(start_duct, doc, view, height_tolerance=0.01):
        """Find all connected ducts/fittings that match shape, size, and z-axis height."""
        run = set()
        to_visit = [start_duct]
        visited = set()
        start_size_obj = Size(str(start_duct.size))

        def shape_key_from_size(size_obj):
            """Create a comparable key from Size using inlet fields only."""
            shape = size_obj.in_shape()
            if shape == "round" and size_obj.in_diameter is not None:
                return ("round", round(size_obj.in_diameter, 2))
            if shape == "oval" and size_obj.in_oval_dia is not None and size_obj.in_oval_flat is not None:
                return ("oval", round(size_obj.in_oval_dia, 2), round(size_obj.in_oval_flat, 2))
            if shape == "rectangle" and size_obj.in_width is not None and size_obj.in_height is not None:
                return ("rect", round(size_obj.in_width, 2), round(size_obj.in_height, 2))
            return ("unknown", str(size_obj.in_size))

        def get_duct_z_coordinate(duct):
            """Extract Z-coordinate (elevation) from centerline, fallback to inlet origin."""
            # Try centerline midpoint Z
            try:
                loc = duct.element.Location
                if hasattr(loc, 'Curve') and loc.Curve:
                    c = loc.Curve
                    p0 = c.GetEndPoint(0)
                    p1 = c.GetEndPoint(1)
                    return (p0.Z + p1.Z) / 2.0
            except Exception:
                pass

            # Fallback to inlet origin Z
            try:
                inlet_data, outlet_data = duct._inlet_outlet_from_revit_xyz()
                if inlet_data and 'origin' in inlet_data:
                    origin = inlet_data['origin']
                    return origin.Z
            except Exception:
                pass
            return None

        start_shape = shape_key_from_size(start_size_obj)
        start_z = get_duct_z_coordinate(start_duct)

        while to_visit:
            duct = to_visit.pop()
            if duct.id in visited:
                continue
            visited.add(duct.id)
            run.add(duct)
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
                        try:
                            connected_duct = RevitDuct(
                                doc, view, connected_elem)
                        except Exception:
                            continue

                        # Check if already visited
                        if connected_duct.id in visited:
                            continue

                        # Parse connected duct shape and size
                        connected_size_obj = Size(str(connected_duct.size))
                        connected_shape = shape_key_from_size(
                            connected_size_obj)

                        # Check Z-axis height difference
                        connected_z = get_duct_z_coordinate(connected_duct)
                        z_difference = abs(
                            connected_z - start_z) if (connected_z is not None and start_z is not None) else None

                        # Match shape, size, and z-axis height
                        if (connected_shape == start_shape and
                            z_difference is not None and
                                z_difference <= height_tolerance):
                            to_visit.append(connected_duct)
        return list(run)

    @staticmethod
    def parse_length_string(length_str):
        """Convert a Revit length string to inches (float)."""
        if not length_str or not isinstance(length_str, str):
            return 0.0

        # Pattern: feet, inches, optional fraction
        pattern = r"(\d+)'\s*-\s*(\d+)?(?:\s+(\d+)/(\d+))?\s*\""
        cleaned = length_str.replace("’", "'").replace(
            "”", '"').replace("″", '"')
        match = re.match(pattern, cleaned)
        if not match:
            # Try to parse as a simple float
            try:
                return float(length_str)
            except Exception:
                return 0.0
        feet = int(match.group(1)) if match.group(1) else 0
        inches = int(match.group(2)) if match.group(2) else 0
        num = int(match.group(3)) if match.group(3) else 0
        denom = int(match.group(4)) if match.group(4) else 1
        fraction = float(num) / float(denom) if denom else 0
        total_inches = feet * 12 + inches + fraction
        return total_inches


# Backward compatibility for older imports/usages.
RrevitRuns = RevitRuns
