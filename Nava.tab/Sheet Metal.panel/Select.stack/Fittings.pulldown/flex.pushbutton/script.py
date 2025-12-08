# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from revit_duct import RevitDuct
from revit_output import print_disclaimer
from pyrevit import revit, script
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, ElementId
from System.Collections.Generic import List


# Button info
# ===================================================
__title__ = "Flex"
__doc__ = """
Find flex longer than 60 inches
"""

# Variables
# ==================================================
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
view = revit.active_view
output = script.get_output()
all_duct = RevitDuct.all(doc, view)

max_lenght = 5.0  # feet

# Class
# =====================================================================
flex_ducts = FilteredElementCollector(doc, view.Id)\
    .OfCategory(BuiltInCategory.OST_FlexDuctCurves)\
    .WhereElementIsNotElementType()\
    .ToElements()

filtered_flex_ducts = []

for d in flex_ducts:
    length_value = d.LookupParameter("Length")
    length = length_value.AsDouble() if length_value else 0
    if length > max_lenght:
        filtered_flex_ducts.append(d)

if filtered_flex_ducts:
    flex_ids = [d.Id for d in filtered_flex_ducts]
    uidoc.Selection.SetElementIds(List[ElementId](flex_ids))

    for i, d in enumerate(filtered_flex_ducts, start=1):
        length_value = d.LookupParameter("Length")
        length = length_value.AsDouble() if length_value else 0
        output.print_md(
            '### No: {:03} | ID: {} | Length {:05.2f}"'.format(
                i,
                output.linkify(d.Id),
                length,
            )
        )

    element_id = [d.Id for d in filtered_flex_ducts]
    output.print_md(
        "# Selected {} flex ducts | {}".format(
            len(filtered_flex_ducts),
            output.linkify(element_id)
        )
    )

    # Final print statements
    print_disclaimer(output)
else:
    output.print_md("No flex ducts found in this view")
