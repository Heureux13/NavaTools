# -*- coding: utf-8 -*-
__title__   = "Testing Class"
__doc__     = """
****************************************************************
Description:
Testing out the class I created
****************************************************************
Author: Jose Nava
"""

# Imports
# ==================================================
from Autodesk.Revit.DB import *
from pyrevit import revit, DB
from revit_duct import RevitDuct
from tag_duct import TagDuct

#.NET Imports
# ==================================================
import clr
clr.AddReference('System')
from System.Collections.Generic import List


# Variables
# ==================================================
app   = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
view = revit.active_view

ducts = (DB.FilteredElementCollector(doc, view.Id)
         .OfCategory(DB.BuiltInCategory.OST_FabricationDuctwork)
         .WhereElementIsNotElementType()
         .ToElements())

# Main Code
# ==================================================

# List every property you want to test
props = [
    "id", "category", "size", "length", "width", "depth",
    "connector_0", "connector_1", "connector_2",
    "system_abbreviation", "family", "double_wall", "insulation",
    "service", "inner_radius",
    "extension_top", "extension_bottom", "extension_right", "extension_left",
    "area", "weight", "angle", "is_full_joint", "insulation_specification",
    "total_weight",
]

print("Testing RevitDuct properties on first 3 elements...")

td = TagDuct(doc, view)
total_tag = td.get_label("jfn_total_weight")

t = DB.Transaction(doc, "Test + Tag ducts")
t.Start()

for d in ducts[:1500]:
    rd = RevitDuct(doc, view, d)
    print("\n--- Element {} ---".format(rd.id))
    for p in props:
        try:
            val = getattr(rd, p)
        except Exception as err:
            val = "[ERROR: {}]".format(err)
        print("{:<20}: {}".format(p, val))

    # --- Tagging section ---
    if rd.total_weight:
        # 1. Write into shared parameter
        p = d.LookupParameter("jfn_total_weight")
        if p:
            if p.StorageType == StorageType.Double:
                p.Set(rd.total_weight)   # direct set for Number parameter
            elif p.StorageType == StorageType.String:
                p.Set(str(rd.total_weight))
    

        # 2. Place tag if not already tagged
        if not td.already_tagged(d, total_tag.Family.Name):
            loc = getattr(d, "Location", None)
            if isinstance(loc, DB.LocationCurve):
                midpt = loc.Curve.Evaluate(0.5, True)
                td.place(d, total_tag, midpt)

    print("rd.weight:", rd.weight)
    print("rd.total_weight:", rd.total_weight)

t.Commit()
