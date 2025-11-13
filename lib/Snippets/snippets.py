# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================
"""⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩"""
# Get XYZ coordinates of selected ducts

# Get selected duct(s)
import Autodesk.Revit.DB as DB
ducts = RevitDuct.from_selection(uidoc, doc)

if not ducts:
    forms.alert("please select one or more duct elements", exitscript=True)

# Header of pop up message
output.print_md("# XYZ of selected ducts")
output.print_md(
    "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++\n")

for duct in ducts:
    xyz = RevitXYZ(duct.element)

    start = xyz.start_point()
    mid = xyz.mid_point()
    end = xyz.end_point()

    # Print duct info
    output.print_md("## Duct ID: {}".format(duct.id))

    if start:
        output.print_md("- Start Point: X: {:.2f}, Y: {:.2f}, Z: {:.2f}".format(
            start.X, start.Y, start.Z
        ))
    if mid:
        output.print_md("- Mid Point: X: {:.2f}, Y: {:.2f}, Z: {:.2f}".format(
            mid.X, mid.Y, mid.Z
        ))
    if end:
        output.print_md("- End Point: X: {:.2f}, Y: {:.2f}, Z: {:.2f}".format(
            end.X, end.Y, end.Z
        ))
    output.print_md("\n---\n")

output.print_md("**Total duct elements processed:** {}".format(len(ducts)))

"""⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧⇧"""

"""⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩⇩"""


def project_to_view(pt, view):
    # For plan views, X and Y are usually the same as model coordinates.
    # For 3D or rotated views, project pt onto the view plane.
    # This example assumes a plan view (Z is up).
    # For more complex views, you'd use view.ViewDirection and view.Origin.
    return pt.X, pt.Y


geo_options = DB.Options()
geo = duct.element.get_Geometry(geo_options)
min_pt = None
min_proj = None

for geom_obj in geo:
    if isinstance(geom_obj, DB.Solid) and geom_obj.Faces.Size > 0:
        for edge in geom_obj.Edges:
            for pt in edge.Tessellate():
                proj_x, proj_y = project_to_view(pt, view)
                if min_proj is None or (
                    proj_x < min_proj[0] or
                    (proj_x == min_proj[0] and proj_y < min_proj[1])
                ):
                    min_proj = (proj_x, proj_y)
                    min_pt = pt

if min_pt:
    output.print_md("- Lower Left Corner (screen/view): X: {:.2f}, Y: {:.2f}, Z: {:.2f}".format(
        min_pt.X, min_pt.Y, min_pt.Z
    ))
