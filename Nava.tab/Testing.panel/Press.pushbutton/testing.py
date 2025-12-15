# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from pyrevit import revit, script, forms
from revit_duct import RevitDuct
from revit_parameter import RevitParameter
from offsets import Offsets
from size import Size
from Autodesk.Revit.DB import Transaction

# Button display information
# =================================================
__title__ = "Press"
__doc__ = """
Assigns offset information about specific duct fittings
"""

# Variables
# ==================================================
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = revit.doc
view = revit.active_view
output = script.get_output()
rp = RevitParameter(doc, app)

fab_ducts = (FilteredElementCollector(doc, view.Id)
             .OfCategory(BuiltInCategory.OST_FabricationDuctwork)
             .WhereElementIsNotElementType()
             .ToElements())

parameter_match = {
    "_duct_offset_center_h": "center_horizontal",
    "_duct_offset_center_v": "center_vertical",
    "_duct_offset_top": "top",
    "_duct_offset_bottom": "bottom",
    "_duct_offset_right": "right",
    "_duct_offset_left": "left",
    "_duct_tag_offset": "tag"
}

family_list = {
    "transition",
    "mitred offset",
    "radius offset",
    "mitered offset",
    "ogee",
    "offset",
    "reducer",
    "square to ø",
}

reducer_square = {
    "transition"
}

reducer_round = {
    "reducer"
}

square_round = {
    "square to ø"
}

offset_list = {
    "ogee",
    "offset",
    "radius offset",
    "mitered offset",
    "mitred offset",
}

if not fab_duct:
    output.print_md("No fab ducts were found")

else:
    try:
        with Transaction(doc, "Set Offsets") as t:
            t.Start()

            for d in fab_ducts:
                tag = None
                fab_ducts = RevitDuct(doc, view, d)
                duct = Size(fab_d.size)
                duct_offsets = Offsets(
                    duct, duct.in_size, duct.out_size, duct.size)
                family = duct_offsets.family.lower().strip()

                if family in family_list:
                    output.print_md(
                        "Family: {} | Length: {:06.2f} | Size: {} | Family: {}".format(
                            output.linkify(d.Id),
                            duct.centerline_length,
                            duct.size,
                            duct.family,
                        ))

                    for param, offset_key in parameter_match.items():
                        value = offset.get(offset_key)
                        if value is not None:
                            rp.set_parameter_value(duct, param, value)
