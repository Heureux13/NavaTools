# -*- coding: utf-8 -*-
"""Filter current selection to non-fabrication pipe (curves and fittings) and reselect."""

from Autodesk.Revit.DB import BuiltInCategory, ElementId
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import revit, script
from revit_duct import RevitDuct
from revit_parameter import RevitParameter
from revit_output import print_disclaimer
from System.Collections.Generic import List

# Button display information
# =================================================
__title__ = "Clean"
__doc__ = """
Takes selected and filters everything out except non-fab pipe
"""

# Variables
# ===================================================
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
view = revit.active_view
output = script.get_output()
ducts = RevitDuct.from_selection(uidoc, doc, view)
rp = RevitParameter(doc, app)

# Helpers
# ===============================================================
def is_non_fab_pipe(el):
    cat = el.Category
    if not cat:
        return False
    return cat.Id in (
        ElementId(BuiltInCategory.OST_PipeCurves),
        ElementId(BuiltInCategory.OST_PipeFitting),
        ElementId(BuiltInCategory.OST_PipeAccessory),
    )

# Main Code
# ==================================================================
sel_ids = list(uidoc.Selection.GetElementIds())


if not sel_ids:
    try:
        picked = uidoc.Selection.PickObjects(
            ObjectType.Element,
            "Select elements to filter to non-fabrication pipe",
        )
        sel_ids = [ref.ElementId for ref in picked]
    except Exception:
        script.exit()

if not sel_ids:
    output.print_md("No elements selected.")
    script.exit()

elements = [doc.GetElement(eid) for eid in sel_ids]
filtered = [e for e in elements if is_non_fab_pipe(e)]

if not filtered:
    output.print_md("No non-fabrication pipe found in selection.")
    script.exit()

duct_run = filtered
out_ids = List[ElementId]()
for e in duct_run:
    out_ids.Add(e.Id)

for i, d in enumerate(duct_run, start=1):
    duct_obj = RevitDuct(doc, view, d)
    family_name = duct_obj.family if duct_obj.family else "Unknown"
    output.print_md(
        "### No: {:03} | ID: {} | Family: {}".format(
            i,
            output.linkify(d.Id),
            family_name
        )
    )

element_ids = [d.Id for d in duct_run]
output.print_md("---")
output.print_md(
    "# Total Elements: {}, {}".format(
        len(duct_run),
        output.linkify(element_ids)
    )
)

# Final print statements
print_disclaimer(output)

uidoc.Selection.SetElementIds(out_ids)
