# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from Autodesk.Revit.DB import FilteredElementCollector, ElementId, CategoryType
from pyrevit import revit, forms, script
from System.Collections.Generic import List
import sys

# Button info
# ===================================================
__title__ = "Hide Annotations"
__doc__ = """
Hides annotation elements by family in the active view.
"""

# Variables
# ==================================================
doc = revit.doc
active_view = doc.ActiveView
output = script.get_output()

# Main Code
# =================================================

# Collect all annotation elements in the active view
annotations_in_view = []
for element in FilteredElementCollector(doc, active_view.Id).WhereElementIsNotElementType():
    if element.Category and element.Category.CategoryType == CategoryType.Annotation:
        annotations_in_view.append(element)

if not annotations_in_view:
    output.print_md('**No annotations found in this view.**')
    sys.exit(0)

# Group annotations by family name
element_by_family = {}
for elem in annotations_in_view:
    try:
        # Get family name or category name as fallback
        family_name = None
        if hasattr(elem, 'Symbol') and elem.Symbol:
            family_name = elem.Symbol.FamilyName
        elif hasattr(elem, 'Name'):
            family_name = elem.Name

        # Use category name if no family name
        if not family_name and elem.Category:
            family_name = elem.Category.Name

        if family_name:
            if family_name not in element_by_family:
                element_by_family[family_name] = []
            element_by_family[family_name].append(elem.Id)
    except Exception:
        pass

if not element_by_family:
    output.print_md('**No annotation families found in this view.**')
    sys.exit(0)

# Create display list with counts
family_display_list = []
for family_name in sorted(element_by_family.keys()):
    count = len(element_by_family[family_name])
    display_name = '{} ({} elements)'.format(family_name, count)
    family_display_list.append(display_name)

# Add "Hide All" option at the top
hide_all_option = '>>> HIDE ALL ANNOTATIONS ({} total) <<<'.format(len(annotations_in_view))
options_list = [hide_all_option] + family_display_list

# Show selection dialog
selected_options = forms.SelectFromList.show(
    options_list,
    title='Select Annotation Families to Hide',
    multiselect=True,
    button_name='Hide Selected'
)

if not selected_options:
    sys.exit(0)

# Collect element IDs for selected families
element_ids = []

if hide_all_option in selected_options:
    # Hide all annotations
    element_ids = [elem.Id for elem in annotations_in_view]
else:
    # Hide selected families
    for option in selected_options:
        # Extract family name from display string (remove count)
        family_name = option.rsplit(' (', 1)[0]
        if family_name in element_by_family:
            element_ids.extend(element_by_family[family_name])

# Apply temporary hide to view
if element_ids:
    id_list = List[ElementId](element_ids)
    with revit.Transaction('Hide Annotations'):
        try:
            active_view.HideElementsTemporary(id_list)
            output.print_md('**Hidden {} annotation elements.**'.format(len(element_ids)))
        except Exception as e:
            output.print_md('**Error: {}**'.format(str(e)))
else:
    output.print_md('**No elements to hide.**')
