# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from pyrevit import revit, forms, script
from Autodesk.Revit.DB import *
from revit_xyz import RevitXYZ

# Button display information
# =================================================
__title__ = "Basic Info"
__doc__ = """
Shows basic information of selected paramters
"""

# Variables
# ==================================================
app = __revit__.Application             # type: Application
uidoc = __revit__.ActiveUIDocument        # type: UIDocument
doc = revit.doc                         # type: Document
view = revit.active_view
output = script.get_output()

show_parameters = {
    'size',
    'item number',
    'specification',
    'insulation specification',
    'fabrication service',
    'mark',
    'type  mark',
    'angle',
    'reference level',
    'upper end top elevation',
    'middle elevation',
    'lower end bottom elevation',
}

# Convert show_parameters to lowercase and strip for comparison
show_parameters_compare = {param.lower().strip() for param in show_parameters}

# Main Code
# ==================================================
selected_ids = uidoc.Selection.GetElementIds()
if not selected_ids:
    forms.alert("please select one or more elements", exitscript=True)

output.print_md("#DUCT INFORMATION")
output.print_md("---")

for elid in selected_ids:
    el = doc.GetElement(elid)
    output.print_md("###---- Parameters for Element ID {} ----".format(el.Id))

    # Collect instance parameters
    param_list = []
    for param in el.Parameters:
        try:
            name = param.Definition.Name
            if name.lower().strip() not in show_parameters_compare:
                continue
            value = param.AsString()
            if value is None:
                value = param.AsValueString()
            if value is None:
                if param.StorageType == StorageType.Double:
                    value = param.AsDouble()
                elif param.StorageType == StorageType.Integer:
                    value = param.AsInteger()
                elif param.StorageType == StorageType.ElementId:
                    value = param.AsElementId()
            param_list.append((name, value, "Instance"))
        except Exception as ex:
            param_list.append((name, "Error - {}".format(ex), "Instance"))

    # Collect type parameters
    elem_type = doc.GetElement(el.GetTypeId())
    if elem_type:
        for param in elem_type.Parameters:
            try:
                name = param.Definition.Name
                if name.lower().strip() not in show_parameters_compare:
                    continue
                value = param.AsString()
                if value is None:
                    value = param.AsValueString()
                if value is None:
                    if param.StorageType == StorageType.Double:
                        value = param.AsDouble()
                    elif param.StorageType == StorageType.Integer:
                        value = param.AsInteger()
                    elif param.StorageType == StorageType.ElementId:
                        value = param.AsElementId()
                param_list.append((name, value, "Type"))
            except Exception as ex:
                param_list.append((name, "Error - {}".format(ex), "Type"))

    # Sort and print
    for name, value, param_type in sorted(param_list, key=lambda x: x[0].lower()):
        display_value = value if value is not None else "None"
        output.print_md(
            "**{}** [{}]:     *{}*".format(name, param_type, display_value))

    # Get XYZ coordinates
    xyz = RevitXYZ(el)

    # Try curve endpoints first
    sp, ep = xyz.curve_endpoints()
    if sp and ep:
        output.print_md("**Location**: Start ({:.3f}, {:.3f}, {:.3f}) → End ({:.3f}, {:.3f}, {:.3f})".format(
            sp.X, sp.Y, sp.Z, ep.X, ep.Y, ep.Z))
    else:
        # Try connector origins
        origins = xyz.connector_origins()
        if origins:
            output.print_md("**Connectors**:")
            for i, o in enumerate(origins):
                output.print_md("  - Connector {}: ({:.3f}, {:.3f}, {:.3f})".format(
                    i, o.X, o.Y, o.Z))
        else:
            output.print_md("**Location**: Not available")

    # Element type information
    try:
        family = el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsString(
        ) if el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM) else "N/A"
        elem_type = el.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsString(
        ) if el.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM) else "N/A"
        output.print_md(
            "**Family**: {} | **Type**: {}".format(family, elem_type))
    except:
        pass

    # Get connected elements via connectors
    try:
        if hasattr(el, 'ConnectorManager'):
            connectors = el.ConnectorManager.Connectors
            if connectors.Size > 0:
                # Collect unique connected elements
                connected_elements = {}
                for connector in connectors:
                    refs = connector.AllRefs
                    for ref in refs:
                        connected_el = ref.Owner
                        connected_id = connected_el.Id.Value
                        if connected_id not in connected_elements:
                            # Try to get item number from connected element
                            item_num_param = connected_el.LookupParameter(
                                'Item Number')
                            item_num = item_num_param.AsString(
                            ) if item_num_param and item_num_param.AsString() else "N/A"
                            connected_elements[connected_id] = item_num

                if connected_elements:
                    output.print_md("**Connected Elements**:")
                    for conn_id, item_num in sorted(connected_elements.items()):
                        output.print_md(
                            "  - ID: {} | Item #: {}".format(conn_id, item_num))
    except Exception as ex:
        pass
