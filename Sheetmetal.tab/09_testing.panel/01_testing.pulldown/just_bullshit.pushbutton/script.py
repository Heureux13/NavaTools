# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import script

# Button info
# ======================================================================
__title__ = 'Bullshit'
__doc__ = '''
Nava's extravaganza nonsense.
'''

# Variables
# ======================================================================

output = script.get_output()



from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB.Structure import *
import clr

doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView

# element you want to tag
element = doc.GetElement(ElementId(123456))  # replace with your element id

t = Transaction(doc, "Curve Tag")
t.Start()

# 1. Get the curve from the element
loc_curve = element.Location
curve = loc_curve.Curve

# 2. Get a Reference to the curve geometry
opt = Options()
geom = element.get_Geometry(opt)

curve_ref = None
for g in geom:
    if isinstance(g, Curve):
        curve_ref = g.Reference
        break

if not curve_ref:
    raise Exception("No curve reference found.")

# 3. Pick a tag type
tag_type_id = doc.GetDefaultElementTypeId(ElementTypeGroup.TagType)

# 4. Create the tag using the curve reference
tag = IndependentTag.Create(
    doc,
    tag_type_id,
    view.Id,
    curve_ref,
    False,
    TagOrientation.Horizontal
)

# 5. Place the tag head somewhere along the curve (midpoint)
midpoint = curve.Evaluate(0.5, True)
tag.TagHeadPosition = midpoint

t.Commit()

