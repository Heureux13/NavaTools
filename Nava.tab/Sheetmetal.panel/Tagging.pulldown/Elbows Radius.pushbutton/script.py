# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from System.Collections.Generic import List
from Autodesk.Revit.ApplicationServices import Application
from Autodesk.Revit.DB import Transaction, Reference, ElementId
from Autodesk.Revit.UI import UIDocument
from pyrevit import revit, forms, DB, script
from revit_element import RevitElement
from revit_duct import RevitDuct, JointSize, DuctAngleAllowance
from revit_xyz import RevitXYZ
from revit_tagging import RevitTagging
import clr

# Button info
# ==================================================
__title__ = "Elbows Radius"
__doc__ = """
************************************************************************
Description:
Select all mitered elbows not 90° and all radius elbows.
************************************************************************
"""

# Variables
# ==================================================
app = __revit__.Application  # type: Application
uidoc = __revit__.ActiveUIDocument  # type: UIDocument
doc = revit.doc  # type: Document
output = script.get_output()
view = revit.active_view
tagger = RevitTagging(doc=doc, view=view)

# Collect ducts in view
# ==================================================
ducts = RevitDuct.all(doc, view)
if not ducts:
    forms.alert("No ducts found in the current view", exitscript=True)

# Duct families
# ==================================================
duct_families = {
    "radius bend": tagger.get_label("_jfn_radius_bend"),
    "elbow": tagger.get_label("_jfn_elbow"),
    "conical tap - wdamper": tagger.get_label("_jfn_conical"),
    "boot tap - wdamper": tagger.get_label("_jfn_boot_tap"),
    "8inch long coupler wdamper": tagger.get_label("_jfn_coupler"),
    "cap": tagger.get_label("_jfn_cap"),
}

# Filter ducts
# ==================================================
rb_ducts = [d for d in ducts if d.family == "Radius Bend"]
e_ducts = [d for d in ducts if d.family == "Elbow"]
conical = [fil_loop == "conicaltap - wdamper"]

fil_ducts = rb_ducts + e_ducts


# Tag Dictionary
# ==================================================
rb_tag = tagger.get_label("_jfn_radius_bend")
e_tag = tagger.get_label("_jfn_elbow")

# Transaction
# ==================================================
t = Transaction(doc, "Radius Elbows Tagging")
t.Start()
try:
    for d in fil_ducts:
        if (tagger.already_tagged(d.element, rad_tag.FamilyName) or
                tagger.already_tagged(d.element, el_tag.FamilyName)):
            continue
        else:
            loc = d.element.location
            if hasattr(loc, "Point") and loc.Point is not None:
                if d.family == "Radius Bend":
                    tagger.place_tag(d.element, rb_tag, loc.Point)
                else:
                    tagger.place_tag(d.element, e_tag, loc.Point)
            elif hasattr(loc, "Curve") and loc.Curve is not None:
                curve = loc.Curve
                midpoint = curve.Evaluate(0.5, True)
                if d.family == "Radius Bend":
                    tagger.place_tag(d.element, rb_tag, midpoint)
                else:
                    tagger.place_tag(d.element, e_tag, midpoint)
            else:
                continue
    t.Commit()
except Exception as e:
    output.print_md("Tag placement error: {}".format(e))
    t.RollBack()
    raise

# Out put results
# ==================================================
output.print_md("## Selected {} short joint(s)".format(len(fil_ducts)))
output.print_md("---")

RevitElement.select_many(uidoc, rad_ducts)
forms.alert("Selected {} radius elbows\nSelected {} Mitered elbows not 90°".format(
    len(rad_ducts), len(el_ducts)))
