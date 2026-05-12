# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import revit, script, DB

# Button info
# ======================================================================
__title__ = 'VAV Clearance'
__doc__ = '''
Find all VAVs in the active view and place FREE FLOATING CLEARANCE on them.
'''

# Variables
# ======================================================================

output = script.get_output()

VAV_KEYWORD = 'vav'
CLEARANCE_FAMILY_NAME = 'FREE FLOATING CLEARANCE'
CLEARANCE_TYPE_NAME = 'FREE FLOATING CLEARANCE'


def _safe_lower(value):
    return (value or '').strip().lower()


def _get_symbol_family_name(symbol):
    try:
        fam = getattr(symbol, 'Family', None)
        if fam:
            return fam.Name
    except Exception:
        pass
    return ''


def _get_symbol_name(symbol):
    try:
        name = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if name:
            return name.AsString() or ''
    except Exception:
        pass
    return ''


def _find_clearance_symbol(doc):
    symbols = (
        DB.FilteredElementCollector(doc)
        .OfClass(DB.FamilySymbol)
        .OfCategory(DB.BuiltInCategory.OST_GenericModel)
        .ToElements()
    )

    exact = None
    fallback = None
    family_needle = _safe_lower(CLEARANCE_FAMILY_NAME)
    type_needle = _safe_lower(CLEARANCE_TYPE_NAME)

    for sym in symbols:
        fam_name = _safe_lower(_get_symbol_family_name(sym))
        typ_name = _safe_lower(_get_symbol_name(sym))
        if fam_name == family_needle and typ_name == type_needle:
            exact = sym
            break
        if family_needle in fam_name and fallback is None:
            fallback = sym

    return exact or fallback


def _as_double_parameter_value(element, names):
    for name in names:
        p = element.LookupParameter(name)
        if p and p.StorageType == DB.StorageType.Double:
            return p.AsDouble()
    return None


def _set_double_parameter(element, name, value):
    if value is None:
        return False
    p = element.LookupParameter(name)
    if not p or p.IsReadOnly or p.StorageType != DB.StorageType.Double:
        return False
    p.Set(value)
    return True


def _get_instance_location_point(element, view):
    loc = element.Location
    if isinstance(loc, DB.LocationPoint):
        return loc.Point
    if isinstance(loc, DB.LocationCurve):
        return loc.Curve.Evaluate(0.5, True)

    bbox = element.get_BoundingBox(view) or element.get_BoundingBox(None)
    if bbox:
        return (bbox.Min + bbox.Max) * 0.5
    return None


def _is_vav_instance(doc, element):
    if not isinstance(element, DB.FamilyInstance):
        return False

    text_parts = []
    try:
        symbol = element.Symbol
        text_parts.append(_get_symbol_family_name(symbol))
        text_parts.append(_get_symbol_name(symbol))
    except Exception:
        pass

    type_el = doc.GetElement(element.GetTypeId())
    if type_el is not None:
        text_parts.append(getattr(type_el, 'FamilyName', ''))
        try:
            text_parts.append(type_el.Name)
        except Exception:
            pass

    combined = _safe_lower(' '.join([t for t in text_parts if t]))
    return VAV_KEYWORD in combined


def _point_key(point, precision=3):
    return (
        round(point.X, precision),
        round(point.Y, precision),
        round(point.Z, precision),
    )


doc = revit.doc
view = doc.ActiveView

vavs = [
    e for e in (
        DB.FilteredElementCollector(doc, view.Id)
        .OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment)
        .WhereElementIsNotElementType()
        .ToElements()
    ) if _is_vav_instance(doc, e)
]

if not vavs:
    output.print_md('No VAV elements were found in the active view.')
    script.exit()

clearance_symbol = _find_clearance_symbol(doc)
if clearance_symbol is None:
    output.print_md('Could not find family/type FREE FLOATING CLEARANCE in Generic Models.')
    script.exit()

existing_clearance_points = set()
for gm in (
        DB.FilteredElementCollector(doc, view.Id)
        .OfCategory(DB.BuiltInCategory.OST_GenericModel)
        .WhereElementIsNotElementType()
        .ToElements()
):
    if not isinstance(gm, DB.FamilyInstance):
        continue
    sym = gm.Symbol
    fam_name = _safe_lower(_get_symbol_family_name(sym))
    if fam_name != _safe_lower(CLEARANCE_FAMILY_NAME):
        continue
    pt = _get_instance_location_point(gm, view)
    if pt:
        existing_clearance_points.add(_point_key(pt))

created = 0
skipped_existing = 0
skipped_no_point = 0
sized_w = 0
sized_l = 0
sized_d = 0

with revit.Transaction('Place VAV Clearance'):
    if not clearance_symbol.IsActive:
        clearance_symbol.Activate()
        doc.Regenerate()

    for vav in vavs:
        point = _get_instance_location_point(vav, view)
        if point is None:
            skipped_no_point += 1
            continue

        key = _point_key(point)
        if key in existing_clearance_points:
            skipped_existing += 1
            continue

        level = doc.GetElement(vav.LevelId) if vav.LevelId != DB.ElementId.InvalidElementId else None
        if level is None and hasattr(view, 'GenLevel'):
            level = view.GenLevel

        instance = None
        if level is not None:
            try:
                instance = doc.Create.NewFamilyInstance(
                    point,
                    clearance_symbol,
                    level,
                    DB.Structure.StructuralType.NonStructural
                )
            except Exception:
                instance = None

        if instance is None:
            instance = doc.Create.NewFamilyInstance(point, clearance_symbol, DB.Structure.StructuralType.NonStructural)

        width = _as_double_parameter_value(vav, ['Equipment Width', 'Width'])
        length = _as_double_parameter_value(vav, ['Equipment Length', 'Length'])
        depth = _as_double_parameter_value(vav, ['Clearance Depth', 'Equipment Height', 'Height'])

        if _set_double_parameter(instance, 'W', width):
            sized_w += 1
        if _set_double_parameter(instance, 'L', length):
            sized_l += 1
        if _set_double_parameter(instance, 'D', depth):
            sized_d += 1

        created += 1
        existing_clearance_points.add(key)

output.print_md('### VAV Clearance Placement Complete')
output.print_md('- VAVs found in active view: **{}**'.format(len(vavs)))
output.print_md('- Clearance instances created: **{}**'.format(created))
output.print_md('- Skipped (already had clearance at same point): **{}**'.format(skipped_existing))
output.print_md('- Skipped (could not determine point): **{}**'.format(skipped_no_point))
output.print_md('- Set parameter W: **{}**'.format(sized_w))
output.print_md('- Set parameter L: **{}**'.format(sized_l))
output.print_md('- Set parameter D: **{}**'.format(sized_d))
