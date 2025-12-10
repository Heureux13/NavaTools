# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from pyrevit import revit, script
from revit_output import print_disclaimer
from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    FabricationPart,
    UnitUtils,
    UnitTypeId,
)

__title__ = "Total Length"
__doc__ = """
Returns total length for selected fabrication pipe(s).
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


def _get_param(el, name, unit=None, as_type="string"):
    """Lightweight parameter getter with unit conversion."""
    p = el.LookupParameter(name)
    if not p:
        return None
    try:
        if as_type == "double":
            val = p.AsDouble()
            if val is None:
                return None
            if unit:
                val = UnitUtils.ConvertFromInternalUnits(val, unit)
            return round(val, 2)
        if as_type == "int":
            return p.AsInteger()
        s = p.AsString()
        return s if s is not None else p.AsValueString()
    except Exception:
        return None


def _get_length_ft(el):
    """Return part length in feet, checking common fabrication parameters."""

    def from_param(name):
        try:
            p = el.LookupParameter(name)
            if p:
                raw = p.AsDouble()
                if raw is not None:
                    return UnitUtils.ConvertFromInternalUnits(raw, UnitTypeId.Feet)
        except Exception:
            return None
        return None

    for pname in ("Length", "NaviateDBS_Length", "Part Length"):
        val = from_param(pname)
        if val is not None:
            return round(val, 2)

    try:
        p = el.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)
        if p:
            raw = p.AsDouble()
            if raw is not None:
                return round(UnitUtils.ConvertFromInternalUnits(raw, UnitTypeId.Feet), 2)
    except Exception:
        pass

    return None


def _get_size(el):
    size = _get_param(el, "Size")
    if size:
        return size
    dia = _get_param(el, "Diameter", unit=UnitTypeId.Inches, as_type="double")
    return dia


def _get_family(el):
    fam = _get_param(el, "NaviateDBS_Family")
    if fam:
        return fam
    fam = _get_param(el, "Family")
    if fam:
        return fam
    try:
        return el.Name
    except Exception:
        return None


def _collect_selected_fab_pipes(uidoc, doc):
    sel_ids = uidoc.Selection.GetElementIds()
    if not sel_ids:
        return []

    cat_id = int(BuiltInCategory.OST_FabricationPipework)
    elements = [doc.GetElement(eid) for eid in sel_ids]

    def is_fab_pipe(el):
        if not isinstance(el, FabricationPart) or not el.Category:
            return False
        cid = el.Category.Id
        try:
            return cid.IntegerValue == cat_id
        except Exception:
            try:
                return cid.Value == cat_id
            except Exception:
                return False

    return [el for el in elements if is_fab_pipe(el)]


# Get selected fabrication pipes
pipes = _collect_selected_fab_pipes(uidoc, doc)

# Select / print loop
if pipes:
    output.print_md('# Selected {} joints of pipe'.format(len(pipes)))
    output.print_md('---')

    # Individual properties
    enriched = []
    for i, el in enumerate(pipes, start=1):

        size = _get_size(el)
        length_ft = _get_length_ft(el)
        family = _get_family(el)
        enriched.append((el, size, length_ft))

        output.print_md(
            '### No: {:03} | ID: {} | Size: {} | Length: {} ft | Family: {}'.format(
                i,
                output.linkify(el.Id),
                size,
                length_ft,
                family
            )
        )

    # Final totals loop and link
    element_ids = [el.Id for el, _, _ in enriched]
    lengths = [length for _, _, length in enriched if length is not None]
    total_length = sum(lengths) if lengths else 0.0

    output.print_md(
        '# Total elements: {} | Total length: {:.2f} ft | {}'.format(
            len(element_ids),
            total_length,
            output.linkify(element_ids)
        )
    )

    # Final print statements
    print_disclaimer(output)
else:
    output.print_md("No pipework found. Select fabrication pipes first.")
