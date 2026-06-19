# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from Autodesk.Revit.DB import ViewDuplicateOption, ViewFamilyType, BuiltInParameter
from Autodesk.Revit.DB import BuiltInCategory
from Autodesk.Revit.DB import View, ElementId, BoundingBoxIntersectsFilter, Outline, View3D, ViewOrientation3D
from Autodesk.Revit.UI import UIView
from pyrevit import script, revit
from Autodesk.Revit.DB import FilteredElementCollector, ViewType, BoundingBoxXYZ, XYZ

# Button info
# ======================================================================
__title__ = 'Create 3D View'
__doc__ = '''
Creats a 3D view of selected elements
'''

# Configuration
# ======================================================================
# Type Name to use for duplicating views (the Type, not the view name)
# This should be the Type Name you want, e.g., "3D View" or "-Working View - Josh"
TEMPLATE_TYPE_NAME = '3D View'

# Variables
# ======================================================================
output = script.get_output()
doc = revit.doc
uidoc = revit.uidoc

# Get Revit username
username = revit.doc.Application.Username
view_name = '3D View - {}'.format(username)

# Get selected elements
selection = uidoc.Selection
selected_ids = selection.GetElementIds()

if selected_ids.Count == 0:
    script.exit()

# Get all selected elements
selected_elements = [doc.GetElement(elem_id) for elem_id in selected_ids]
selected_elements = [elem for elem in selected_elements if elem is not None]

if not selected_elements:
    script.exit()

# Get source view before deleting anything
source_view = uidoc.ActiveView
if source_view is None or source_view.ViewType != ViewType.ThreeD:
    for view in FilteredElementCollector(doc).OfClass(View):
        if view.ViewType == ViewType.ThreeD and not view.IsTemplate:
            source_view = view
            break

if source_view is None:
    script.exit()


# Find the target view type
target_type_id = None
for vft in FilteredElementCollector(doc).OfClass(ViewFamilyType):
    try:
        param = vft.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if param is not None and param.AsString() == TEMPLATE_TYPE_NAME:
            target_type_id = vft.Id
            break
    except:
        pass

# Find existing view and ensure source is not the view we're deleting
existing_view_id = None
for view in FilteredElementCollector(doc).OfClass(View):
    if view.Name == view_name and view.ViewType == ViewType.ThreeD:
        existing_view_id = view.Id
        if view.Id == source_view.Id:
            # Source IS the target view - pick a different source
            for v in FilteredElementCollector(doc).OfClass(View):
                if v.ViewType == ViewType.ThreeD and not v.IsTemplate and v.Id != view.Id:
                    source_view = v
                    break
        break

# Delete old view in its own transaction
if existing_view_id is not None:
    with revit.Transaction('Delete Old 3D View'):
        doc.Delete(existing_view_id)

# Create new view by duplicating source
with revit.Transaction('Create 3D View'):
    view_id = source_view.Duplicate(ViewDuplicateOption.Duplicate)
    view_3d = doc.GetElement(view_id)
    view_3d.Name = view_name
    if target_type_id is not None:
        view_3d.ChangeTypeId(target_type_id)


# Get combined bounding box of all selected elements
combined_min = None
combined_max = None

for element in selected_elements:
    # Try to get bounding box without specifying a view
    try:
        bbox = element.get_BoundingBox(None)
        if bbox is not None:
            if combined_min is None:
                combined_min = bbox.Min
                combined_max = bbox.Max
            else:
                # Expand combined bbox to include this element
                combined_min = XYZ(
                    min(combined_min.X, bbox.Min.X),
                    min(combined_min.Y, bbox.Min.Y),
                    min(combined_min.Z, bbox.Min.Z)
                )
                combined_max = XYZ(
                    max(combined_max.X, bbox.Max.X),
                    max(combined_max.Y, bbox.Max.Y),
                    max(combined_max.Z, bbox.Max.Z)
                )
    except:
        pass

if combined_min is None:
    output.print_md(
        '**Error:** Could not get bounding box - all elements returned None')
    script.exit()

output.print_md('BBox Min: ({:.2f}, {:.2f}, {:.2f})'.format(
    combined_min.X, combined_min.Y, combined_min.Z))
output.print_md('BBox Max: ({:.2f}, {:.2f}, {:.2f})'.format(
    combined_max.X, combined_max.Y, combined_max.Z))

# Expand bounding box by 2 feet (24 inches)
expansion = 2.0  # feet, Revit API uses feet by default

min_pt = combined_min
max_pt = combined_max

expanded_min = XYZ(min_pt.X - expansion, min_pt.Y -
                   expansion, min_pt.Z - expansion)
expanded_max = XYZ(max_pt.X + expansion, max_pt.Y +
                   expansion, max_pt.Z + expansion)

# Create new section box
section_box = BoundingBoxXYZ()
section_box.Min = expanded_min
section_box.Max = expanded_max

# Apply section box and enable MEP Fabrication visibility inside a transaction

with revit.Transaction('Set View'):
    view_3d.SetSectionBox(section_box)

    # Force-enable MEP Fabrication Part categories visibility
    mep_cats = [
        BuiltInCategory.OST_FabricationDuctwork,
        BuiltInCategory.OST_FabricationPipework,
        BuiltInCategory.OST_FabricationContainment,
        BuiltInCategory.OST_FabricationHangers,
    ]
    for bic in mep_cats:
        try:
            cat = doc.Settings.Categories.get_Item(bic)
            if cat is not None:
                view_3d.SetCategoryHidden(cat.Id, False)
        except:
            pass

# Verify section box was applied
applied_box = view_3d.GetSectionBox()
output.print_md('Section box applied - Min: ({:.2f}, {:.2f}, {:.2f}) Max: ({:.2f}, {:.2f}, {:.2f})'.format(
    applied_box.Min.X, applied_box.Min.Y, applied_box.Min.Z,
    applied_box.Max.X, applied_box.Max.Y, applied_box.Max.Z))

# Switch to the view AFTER transaction is closed
uidoc.ActiveView = view_3d

# Zoom to fit - find the specific UIView for this view
try:
    ui_views = uidoc.GetOpenUIViews()
    active_ui_view = None
    for uiv in ui_views:
        if uiv.ViewId == view_3d.Id:
            active_ui_view = uiv
            break
    if active_ui_view:
        active_ui_view.ZoomToFit()
        output.print_md(
            'ZoomToFit called on view: **{}**'.format(view_3d.Name))
    else:
        output.print_md('**Warning:** Could not find UIView for this view')
except Exception as e:
    output.print_md('Zoom error: {}'.format(str(e)))
