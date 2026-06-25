# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import revit, script
from System.Collections.Generic import List
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ElementId
)

# Button info
# ======================================================================
__title__ = 'Section Cover'
__doc__ = '''
five flying fly fleets feeble feet
'''

# Variables
# ======================================================================

doc = revit.doc  # type: ignore[attr-defined]
uidoc = revit.uidoc  # type: ignore[attr-defined]
output = script.get_output()

parameter = 'view name'
param_value = {
    'horizontal',
    'vertical',
}


def get_sections_on_view(doc, plan_view):
    sections = []
    collector = (
        FilteredElementCollector(doc, plan_view.Id)
        .WhereElementIsNotElementType()
    )
    for elem in collector:
        try:
            cat = elem.Category
            if not cat or cat.Name != 'Views':
                continue
        except Exception:
            continue

        # Confirm it is a section by checking name/family tokens
        tokens = []
        for attr in ('Name',):
            try:
                tokens.append((getattr(elem, attr, '') or '').lower())
            except Exception:
                pass
        for pname in ('Family', 'Family and Type', 'Type'):
            try:
                p = elem.LookupParameter(pname)
                if p:
                    tokens.append(
                        (p.AsString() or p.AsValueString() or '').lower())
            except Exception:
                pass

        if 'section' in ' '.join(tokens):
            sections.append(elem)

    return sections


def _get_param_value(elem, param_name):
    target = param_name.strip().lower()
    try:
        for p in elem.Parameters:
            try:
                if p.Definition.Name.strip().lower() == target:
                    return (p.AsString() or p.AsValueString() or '').strip().lower()
            except Exception:
                continue
    except Exception:
        pass
    return ''


def hide_sections_except(doc, plan_view, keep_param, keep_values):
    sections = get_sections_on_view(doc, plan_view)
    ids_to_hide = List[ElementId]()

    for sec in sections:
        value = _get_param_value(sec, keep_param)
        if value not in keep_values:
            ids_to_hide.Add(sec.Id)

    if not ids_to_hide.Count:
        output.print_md('No sections to hide.')
        return 0

    with revit.Transaction('Temporarily hide section markers'):
        plan_view.HideElementsTemporary(ids_to_hide)

    return ids_to_hide.Count


hidden_count = hide_sections_except(
    doc, doc.ActiveView, parameter, {v.lower() for v in param_value})
# output.print_md('Temporarily hid {} section(s).'.format(hidden_count))

# output.print_md('Testing script is running.')
