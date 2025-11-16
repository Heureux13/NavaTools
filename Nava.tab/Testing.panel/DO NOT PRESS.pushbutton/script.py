# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
import sys
from Autodesk.Revit.DB import Transaction
from System.Collections.Generic import List
from revit_parameter import RevitParameter
from revit_element import RevitElement
from tag_duct import TagDuct
from revit_duct import RevitDuct, JointSize, CONNECTOR_THRESHOLDS
from revit_xyz import RevitXYZ
from Autodesk.Revit.ApplicationServices import Application
from Autodesk.Revit.UI import UIDocument
from pyrevit import revit, script, forms, DB
from Autodesk.Revit.DB import *
import clr

# Button display information
# =================================================
__title__ = "DO NOT PRESS"
__doc__ = """******************************************************************
Description:

Current goal fucntion of button is: select only spiral duct.

++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
This button is for testin code snippets and ideas before implementing them.
Odds are it will be constantly changing and not useful, its entire purpose
is for the author to have a quick button to test whatever code they are working on.
If you press it could do nothing, throw an error, or change something in your model.
Once working, it will most likely be moved to a more permanent location.
******************************************************************"""

# Variables
# ==================================================
app = __revit__.Application             # type: Application
uidoc = __revit__.ActiveUIDocument        # type: UIDocument
doc = revit.doc                         # type: Document
view = revit.active_view
output = script.get_output()

# CONFIG: set your parameter names here (instance parameters on Fabrication Parts)
PARM_FO_CLASS = "FO_Class"      # e.g., "FOT", "FOB", "CL" + optional side arrow
PARM_FO_VERT = "FO_Vertical"   # e.g., "↑2\"" or "↓1-1/2\""
PARM_FO_SIDE = "FO_Side"       # e.g., "→2\"" or "←1\""
# optional numeric length param (internal feet)
PARM_FO_VSHIFT = "FO_VertShift"
# optional numeric length param (internal feet)
PARM_FO_SSHIFT = "FO_SideShift"


def set_text_param(el, name, value):
    p = el.LookupParameter(name)
    if p and not p.IsReadOnly:
        p.Set(value or "")
        return True
    return False


def set_len_param(el, name, feet_value_abs):
    # Write a length value (in feet) into a length parameter (internal units = feet)
    p = el.LookupParameter(name)
    if p and not p.IsReadOnly and feet_value_abs is not None:
        try:
            val_internal = UnitUtils.ConvertToInternalUnits(
                feet_value_abs, UnitTypeId.Feet)
            p.Set(val_internal)
            return True
        except Exception as e:
            output.print_md("Could not set {}: {}".format(name, e))
    return False


# Get selection or prompt
sel_ids = list(uidoc.Selection.GetElementIds())
if not sel_ids:
    forms.alert(
        "Select one or more Fabrication transitions/reducers and re-run.", exitscript=True)

parts = []
for eid in sel_ids:
    el = doc.GetElement(eid)
    if isinstance(el, DB.FabricationPart):
        parts.append(el)

if not parts:
    forms.alert("No FabricationPart elements in selection.", exitscript=True)

t = Transaction(doc, "Annotate Transitions (FOT/FOB/CL/FOS)")
t.Start()
try:
    for p in parts:
        rxyz = RevitXYZ(p)
        # Force fresh import
        import sys
        if 'revit_xyz' in sys.modules:
            reload(sys.modules['revit_xyz'])
            from revit_xyz import RevitXYZ
            rxyz = RevitXYZ(p)
        info = rxyz.analyze_transition()

        # Compose strings
        vclass = info.get('vertical_class') or ""
        # "↑2\"" or "↓1-1/2\"" or ""
        varrow = info.get('vertical_arrow') or ""
        sclass = info.get('side_class') or ""       # "FOS" or ""
        s_arrow = info.get('side_arrow') or ""      # "→2"" or "←1"" or ""

        # Main classification string: vclass plus side arrow if present
        main = vclass
        if s_arrow:
            main = (main + " · " + s_arrow) if main else s_arrow

        # Write text parameters (ignore if missing/read-only)
        if PARM_FO_CLASS:
            set_text_param(p, PARM_FO_CLASS, main)
        if PARM_FO_VERT:
            set_text_param(p, PARM_FO_VERT,  varrow)
        if PARM_FO_SIDE:
            set_text_param(p, PARM_FO_SIDE,  s_arrow)

        # Optional: also store numeric offsets as absolute values (feet)
        vshift = abs(info.get('vertical_shift_ft') or 0.0)
        sshift = abs(info.get('side_shift_ft') or 0.0)
        if PARM_FO_VSHIFT:
            set_len_param(p, PARM_FO_VSHIFT, vshift)
        if PARM_FO_SSHIFT:
            set_len_param(p, PARM_FO_SSHIFT, sshift)

        # Debug to console
        output.print_md("- Element {} → {} | {} | {}".format(
            p.Id, main, varrow, s_arrow
        ))
        dbg = info.get('debug', {})
        output.print_md("Debug: vclass={} vshift={:.3f}ft top_diff={:.3f}ft bot_diff={:.3f}ft side={:.3f}ft".format(
            vclass or 'None',
            info.get('vertical_shift_ft') or 0.0,
            dbg.get('top_diff_ft') or 0.0,
            dbg.get('bot_diff_ft') or 0.0,
            info.get('side_shift_ft') or 0.0
        ))
        if dbg.get('reason'):
            output.print_md("  Reason: {} | num_faces={}".format(
                dbg.get('reason'), dbg.get('num_faces', 0)))

    t.Commit()
    output.print_md("Done: updated {} parts.".format(len(parts)))
except Exception as e:
    t.RollBack()
    output.print_md("Update failed: {}".format(e))
    raise

# -*- coding: utf-8 -*-

doc = revit.doc
uidoc = __revit__.ActiveUIDocument
output = script.get_output()

PARAM_NAME = "_jfn_offset"  # your target text parameter


def set_text(el, name, value):
    p = el.LookupParameter(name)
    if p and not p.IsReadOnly:
        p.Set(value or "")
        return True
    return False


# Use current selection
sel_ids = list(uidoc.Selection.GetElementIds())
if not sel_ids:
    forms.alert(
        "Select one or more Fabrication transitions/reducers and re-run.", exitscript=True)

parts = []
for eid in sel_ids:
    el = doc.GetElement(eid)
    if isinstance(el, DB.FabricationPart):
        parts.append(el)

if not parts:
    forms.alert("No FabricationPart elements in selection.", exitscript=True)

t = Transaction(doc, "Set _jfn_offset")
t.Start()
try:
    for p in parts:
        info = RevitXYZ(p).analyze_transition()

        # Build what you want to display in _jfn_offset
        vclass = info.get('vertical_class') or ""   # FOT | FOB | CL | ""
        s_arrow = info.get('side_arrow') or ""      # →2" | ←1-1/2" | ""
        v_arrow = info.get('vertical_arrow') or ""  # ↑1" | ↓3/4" | ""

        # Improved: always include non-empty components; if ALL empty, fallback to "CL?" or "NO-OFFSET"
        if not any([vclass, s_arrow, v_arrow]):
            label = "CL"
        else:
            label = " · ".join([p for p in [vclass, s_arrow, v_arrow] if p])

        if not set_text(p, PARAM_NAME, label):
            output.print_md(
                "Could not set {} on element {}".format(PARAM_NAME, p.Id))
        else:
            output.print_md("- {} → {}".format(p.Id, label))

    t.Commit()
    output.print_md("Updated {} part(s).".format(len(parts)))
except Exception as e:
    t.RollBack()
    output.print_md("Failed: {}".format(e))
    raise
