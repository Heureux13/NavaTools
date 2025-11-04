"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""


from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import FabricationPart
from revit_duct import RevitDuct

uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document
view  = uidoc.ActiveView

for eid in uidoc.Selection.GetElementIds():
    el = doc.GetElement(eid)
    if not el or not isinstance(el, FabricationPart):
        continue

    duct = RevitDuct(doc, view, el)
    print("ID:", duct.id, "Category:", duct.category, "Size:", duct.size)
    print("ID:", duct.id, "Category:", duct.category, "Length:", duct.length)
    print("ID:", duct.id, "Category:", duct.category, "connector_0:", duct.connector_0)
    print("ID:", duct.id, "Category:", duct.category, "connector_1:", duct.connector_1)
    print("ID:", duct.id, "Category:", duct.category, "connector_2:", duct.connector_2)
    print("ID:", duct.id, "Category:", duct.category, "width:", duct.width)
    print("ID:", duct.id, "Category:", duct.category, "depth:", duct.depth)
    print("ID:", duct.id, "Category:", duct.category, "family:", duct.family)
    print("ID:", duct.id, "Category:", duct.category, "double_wall:", duct.double_wall)
    print("ID:", duct.id, "Category:", duct.category, "insulation:", duct.insulation)
    print("ID:", duct.id, "Category:", duct.category, "service:", duct.service)
    print("ID:", duct.id, "Category:", duct.category, "inner_radius:", duct.inner_radius)
    print("ID:", duct.id, "Category:", duct.category, "extension_top:", duct.extension_top)
    print("ID:", duct.id, "Category:", duct.category, "extension_bottom:", duct.extension_bottom)
    print("ID:", duct.id, "Category:", duct.category, "extension_right:", duct.extension_right)
    print("ID:", duct.id, "Category:", duct.category, "extension_left:", duct.extension_left)
    print("ID:", duct.id, "Category:", duct.category, "angle:", duct.angle)
    print("ID:", duct.id, "Category:", duct.category, "is_full_joint:", duct.is_full_joint)