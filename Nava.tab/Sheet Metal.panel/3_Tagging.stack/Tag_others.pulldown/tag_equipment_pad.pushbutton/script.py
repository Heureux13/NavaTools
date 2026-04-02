# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import script
from pyrevit import revit
from Autodesk.Revit.DB import (
    BuiltInCategory,
    FamilySymbol,
    FilteredElementCollector,
    IndependentTag,
    Transaction,
    XYZ,
)
from tagging.revit_tagging import RevitTagging  # type: ignore[reportMissingImports]

# Button info
# ======================================================================
__title__ = 'Tag Equipment Pads'
__doc__ = '''
Tags only equipment pads in the active view with _umi_equi_pad and updates value from equipment pad height to _make.
'''

# Variables
# ======================================================================

output = script.get_output()
doc = revit.doc
view = revit.active_view
tagger = RevitTagging(doc, view)

TAG_NAME = '_umi_equi_pad'
TARGET_CATEGORY = BuiltInCategory.OST_MechanicalEquipment


def _norm(text):
    return (text or '').strip().lower()


def _find_equipment_tag_symbol(doc, target_name):
    needle = _norm(target_name)
    if not needle:
        return None

    symbols = []
    for bic in (
            BuiltInCategory.OST_MechanicalEquipmentTags,
            BuiltInCategory.OST_MultiCategoryTags,
    ):
        symbols.extend(
            FilteredElementCollector(doc)
            .OfCategory(bic)
            .OfClass(FamilySymbol)
            .ToElements()
        )

    exact_matches = []
    contains_matches = []
    for sym in symbols:
        fam = getattr(sym, 'Family', None)
        fam_name = fam.Name if fam else ''
        type_name = getattr(sym, 'Name', '') or ''
        fam_norm = _norm(fam_name)
        type_norm = _norm(type_name)
        label = _norm('{} {}'.format(fam_name, type_name))

        if needle == fam_norm or needle == type_norm:
            exact_matches.append(sym)
        elif needle in label:
            contains_matches.append(sym)

    if exact_matches:
        return exact_matches[0]
    if contains_matches:
        return contains_matches[0]
    return None


def _is_equipment_pad(elem):
    if not elem or not elem.Category:
        return False

    try:
        if elem.Category.Id.IntegerValue != int(TARGET_CATEGORY):
            return False
    except Exception:
        return False

    fam_name = ''
    type_name = ''
    try:
        symbol = getattr(elem, 'Symbol', None)
        fam = getattr(symbol, 'Family', None)
        fam_name = fam.Name if fam else ''
        type_name = getattr(symbol, 'Name', '') or ''
    except Exception:
        pass

    fam_norm = _norm(fam_name)
    type_norm = _norm(type_name)

    # Keep matching strict to avoid tagging non-pad equipment.
    return (
        fam_norm == 'me_equipment pad'
        or 'equipment pad' in fam_norm
        or 'equipment pad' in type_norm
    )


def _center_point(elem, view):
    try:
        loc = elem.Location
        if hasattr(loc, 'Point') and loc.Point:
            return loc.Point
    except Exception:
        pass

    bbox = elem.get_BoundingBox(view) if view else None
    if not bbox:
        return None

    min_pt = bbox.Min
    max_pt = bbox.Max
    return XYZ(
        (min_pt.X + max_pt.X) / 2.0,
        (min_pt.Y + max_pt.Y) / 2.0,
        (min_pt.Z + max_pt.Z) / 2.0,
    )


def _build_existing_tag_map(view, tag_symbol_id):
    existing_tag_map = {}
    existing_tags = (
        FilteredElementCollector(doc, view.Id)
        .OfClass(IndependentTag)
        .ToElements()
    )

    for tag in existing_tags:
        try:
            if tag.GetTypeId() != tag_symbol_id:
                continue

            tagged_ids = []
            try:
                tagged_ids = list(tag.GetTaggedLocalElementIds() or [])
            except Exception:
                pass

            if not tagged_ids:
                try:
                    tid = tag.TaggedLocalElementId
                    if tid:
                        tagged_ids = [tid]
                except Exception:
                    pass

            for tid in tagged_ids:
                try:
                    existing_tag_map[tid.IntegerValue] = True
                except Exception:
                    continue
        except Exception:
            continue

    return existing_tag_map


def _pad_make_value_from_height(elem):
    height_param = elem.LookupParameter('Equipment Pad Height')
    if not height_param:
        return None

    # Prefer numeric conversion so formatting is consistent (e.g. 06" PAD).
    try:
        height_ft = height_param.AsDouble()
        height_in = int(round(height_ft * 12.0))
        if height_in <= 0:
            return None
        return '{:02d}" PAD'.format(height_in)
    except Exception:
        pass

    try:
        raw = (height_param.AsValueString() or '').strip()
        if not raw:
            return None
        return '{} PAD'.format(raw)
    except Exception:
        return None


def _sync_make_from_height(elem):
    make_param = elem.LookupParameter('_make')
    if not make_param or make_param.IsReadOnly:
        return False, 'Missing or read-only _make'

    make_value = _pad_make_value_from_height(elem)
    if not make_value:
        return False, 'Missing Equipment Pad Height value'

    current_value = (make_param.AsString() or '').strip()
    if current_value == make_value:
        return False, ''

    make_param.Set(make_value)
    return True, ''


tag_symbol = _find_equipment_tag_symbol(doc, TAG_NAME)
if not tag_symbol:
    output.print_md('## Missing tag family/type: {}'.format(TAG_NAME))
    script.exit()

equipment_in_view = (
    FilteredElementCollector(doc, view.Id)
    .OfCategory(TARGET_CATEGORY)
    .WhereElementIsNotElementType()
    .ToElements()
)

pads = [e for e in equipment_in_view if _is_equipment_pad(e)]

if not pads:
    output.print_md('## No equipment pads found in the active view.')
    script.exit()

existing_tag_map = _build_existing_tag_map(view, tag_symbol.Id)

placed = []
already_tagged = []
failed = []
make_updated = []
make_failed = []

t = Transaction(doc, 'Tag Equipment Pads')
t.Start()
try:
    for elem in pads:
        try:
            elem_id = elem.Id.IntegerValue

            make_changed, make_reason = _sync_make_from_height(elem)
            if make_changed:
                make_updated.append(elem)
            elif make_reason:
                make_failed.append((elem, make_reason))

            if existing_tag_map.get(elem_id):
                already_tagged.append(elem)
                continue

            pt = _center_point(elem, view)
            if not pt:
                failed.append((elem, 'Unable to determine tag point'))
                continue

            tagger.place_tag(elem, tag_symbol, pt)
            placed.append(elem)
            existing_tag_map[elem_id] = True
        except Exception as ex:
            failed.append((elem, str(ex)))

    t.Commit()
except Exception:
    t.RollBack()
    raise

output.print_md(
    '## Equipment Pad Tagging: placed {}, already tagged {}, failed {}'.format(
        len(placed),
        len(already_tagged),
        len(failed),
    )
)

output.print_md(
    '## _make Sync: updated {}, issues {}'.format(
        len(make_updated),
        len(make_failed),
    )
)

if failed:
    output.print_md('### Failed Elements')
    for elem, reason in failed:
        output.print_md('- ID {}: {}'.format(output.linkify(elem.Id), reason))

if make_failed:
    output.print_md('### _make Sync Issues')
    for elem, reason in make_failed:
        output.print_md('- ID {}: {}'.format(output.linkify(elem.Id), reason))
