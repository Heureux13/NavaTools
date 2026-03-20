# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from revit_element import RevitElement
from pyrevit import revit, script
from Autodesk.Revit.DB import (
    BuiltInCategory,
    FabricationPart,
    FilteredElementCollector,
    StorageType,
)

# Button info
# ===================================================
__title__ = "Skip GRD Duct"
__doc__ = """
Find GRD-connected ductwork (excluding work tap families)
and set Item Number to "skip".
"""

# Variables
# ==================================================
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
view = revit.active_view
output = script.get_output()

# Main Code
# ==================================================

SKIP_VALUE = "skip"
ITEM_NUMBER_PARAM = "item number"
EXCLUDED_FAMILY_TOKEN = "work tap"
MIN_VERTICAL_Z_COMPONENT = 0.95


def _iter_connectors(element):
    """Yield connectors from an element across common connector access paths."""
    connector_sets = []

    connector_manager = getattr(element, "ConnectorManager", None)
    if connector_manager is not None:
        connector_sets.append(getattr(connector_manager, "Connectors", None))

    mep_model = getattr(element, "MEPModel", None)
    if mep_model is not None:
        mep_connector_manager = getattr(mep_model, "ConnectorManager", None)
        if mep_connector_manager is not None:
            connector_sets.append(getattr(mep_connector_manager, "Connectors", None))

    get_connectors = getattr(element, "GetConnectors", None)
    if callable(get_connectors):
        try:
            connector_sets.append(get_connectors())
        except Exception:
            pass

    for connector_set in connector_sets:
        if connector_set is None:
            continue
        try:
            for connector in connector_set:
                if connector is not None:
                    yield connector
        except Exception:
            continue


def _connected_owners(element):
    """Return unique owner elements directly connected to this element."""
    owners = {}
    for connector in _iter_connectors(element):
        if not getattr(connector, "IsConnected", False):
            continue

        try:
            refs = list(connector.AllRefs)
        except Exception:
            refs = []

        for ref in refs:
            owner = getattr(ref, "Owner", None)
            if owner is None:
                continue
            if owner.Id == element.Id:
                continue
            owners[owner.Id.IntegerValue] = owner

    return owners.values()


def _family_name(element):
    family_name = ""
    try:
        symbol = getattr(element, "Symbol", None)
        family = getattr(symbol, "Family", None) if symbol else None
        family_name = (getattr(family, "Name", "") or "").strip()
    except Exception:
        family_name = ""

    if family_name:
        return family_name

    type_id = element.GetTypeId()
    if type_id:
        type_elem = doc.GetElement(type_id)
        if type_elem and hasattr(type_elem, "FamilyName"):
            return (type_elem.FamilyName or "").strip()

    return ""


def _set_item_number_skip(element):
    for param in element.Parameters:
        try:
            param_name = (param.Definition.Name or "").strip().lower()
        except Exception:
            continue

        if param_name != ITEM_NUMBER_PARAM:
            continue
        if param.IsReadOnly:
            return False, "read-only"
        if param.StorageType != StorageType.String:
            return False, "non-string"

        try:
            current_val = param.AsString()
            if (current_val or "").strip().lower() == SKIP_VALUE:
                return True, "already"

            param.Set(SKIP_VALUE)
            return True, "updated"
        except Exception:
            return False, "set-failed"

    return False, "missing"


def _is_vertical_or_near_vertical(element):
    """True when the duct axis is vertical or close to vertical."""
    # Preferred: use location curve direction.
    try:
        location = getattr(element, "Location", None)
        curve = getattr(location, "Curve", None)
        if curve is not None:
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            axis = p1 - p0
            length = axis.GetLength()
            if length and length > 1e-9:
                z_component = abs(axis.Z) / float(length)
                return z_component >= MIN_VERTICAL_Z_COMPONENT
    except Exception:
        pass

    # Fallback: derive axis from farthest connector pair.
    origins = []
    for connector in _iter_connectors(element):
        try:
            origin = connector.Origin
            if origin is not None:
                origins.append(origin)
        except Exception:
            continue

    if len(origins) < 2:
        return False

    max_len = 0.0
    best_vec = None
    for i in range(len(origins)):
        for j in range(i + 1, len(origins)):
            vec = origins[j] - origins[i]
            seg_len = vec.GetLength()
            if seg_len > max_len:
                max_len = seg_len
                best_vec = vec

    if best_vec is None or max_len <= 1e-9:
        return False

    z_component = abs(best_vec.Z) / float(max_len)
    return z_component >= MIN_VERTICAL_Z_COMPONENT


# 1) Collect all GRDs (duct terminals) in active view and select them.
grds = list(
    FilteredElementCollector(doc, view.Id)
    .OfCategory(BuiltInCategory.OST_DuctTerminal)
    .WhereElementIsNotElementType()
    .ToElements()
)

if not grds:
    output.print_md("## No GRDs found in this view.")
    script.exit()

RevitElement.select_many(uidoc, grds)

# 2) Find connected ductwork and exclude work tap families.
target_ducts_by_id = {}
non_vertical_count = 0
for grd in grds:
    for connected in _connected_owners(grd):
        if not isinstance(connected, FabricationPart):
            continue

        category = getattr(connected, "Category", None)
        if category is None or category.Id.IntegerValue != int(BuiltInCategory.OST_FabricationDuctwork):
            continue

        family_name = _family_name(connected).lower()
        if EXCLUDED_FAMILY_TOKEN in family_name:
            continue

        if not _is_vertical_or_near_vertical(connected):
            non_vertical_count += 1
            continue

        target_ducts_by_id[connected.Id.IntegerValue] = connected

target_ducts = list(target_ducts_by_id.values())
if not target_ducts:
    output.print_md("## Selected {} GRDs, but no connected non-work-tap ductwork was found.".format(len(grds)))
    script.exit()

# 3) Set Item Number = "skip".
updated = []
already = []
missing = []
readonly_or_invalid = []

with revit.Transaction('Set Item Number to skip for GRD-connected ductwork'):
    for duct in target_ducts:
        ok, reason = _set_item_number_skip(duct)
        if ok and reason == "updated":
            updated.append(duct.Id)
        elif ok and reason == "already":
            already.append(duct.Id)
        elif reason == "missing":
            missing.append(duct.Id)
        else:
            readonly_or_invalid.append(duct.Id)

output.print_md("## Selected {} GRDs in active view.".format(len(grds)))
output.print_md("## Connected non-work-tap ductwork found: {}".format(len(target_ducts)))
output.print_md("## Connected ducts skipped for non-vertical orientation: {}".format(non_vertical_count))
output.print_md("## Item Number set to skip: {}".format(len(updated)))

if already:
    output.print_md("Already skip: {}".format(len(already)))
if missing:
    output.print_md("Missing Item Number parameter: {}".format(len(missing)))
if readonly_or_invalid:
    output.print_md("Read-only/non-string/failed to set: {}".format(len(readonly_or_invalid)))
