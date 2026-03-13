# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
from pyrevit import revit, forms
from Autodesk.Revit.DB import BuiltInCategory, FilteredElementCollector, ElementId
from System.Collections.Generic import List
import sys

# Button info
# ===================================================
__title__ = "Isolate by Selection"
__doc__ = """
Isolates the active view to only selected categories.
"""

# Variables
# ==================================================
doc = revit.doc
active_view = doc.ActiveView

# Flat list of categories with friendly names
category_options = {
    'Ducts': BuiltInCategory.OST_DuctCurves,
    'Flex Ducts': BuiltInCategory.OST_FlexDuctCurves,
    'Duct Fittings': BuiltInCategory.OST_DuctFitting,
    'Duct Accessories': BuiltInCategory.OST_DuctAccessory,
    'Duct Terminals': BuiltInCategory.OST_DuctTerminal,
    'Mechanical Equipment': BuiltInCategory.OST_MechanicalEquipment,
    'Duct Insulation': BuiltInCategory.OST_DuctInsulations,
    'Fabrication Ductwork': BuiltInCategory.OST_FabricationDuctwork,
    'Fabrication Hangers': BuiltInCategory.OST_FabricationHangers,

    'Pipes': BuiltInCategory.OST_PipeCurves,
    'Flex Pipes': BuiltInCategory.OST_FlexPipeCurves,
    'Pipe Fittings': BuiltInCategory.OST_PipeFitting,
    'Pipe Accessories': BuiltInCategory.OST_PipeAccessory,
    'Sprinklers': BuiltInCategory.OST_Sprinklers,
    'Plumbing Fixtures': BuiltInCategory.OST_PlumbingFixtures,
    'Pipe Insulation': BuiltInCategory.OST_PipeInsulations,
    'Fabrication Pipework': BuiltInCategory.OST_FabricationPipework,

    'Electrical Equipment': BuiltInCategory.OST_ElectricalEquipment,
    'Electrical Fixtures': BuiltInCategory.OST_ElectricalFixtures,
    'Lighting Fixtures': BuiltInCategory.OST_LightingFixtures,
    'Conduit': BuiltInCategory.OST_Conduit,
    'Conduit Fittings': BuiltInCategory.OST_ConduitFitting,
    'Cable Trays': BuiltInCategory.OST_CableTray,
    'Cable Tray Fittings': BuiltInCategory.OST_CableTrayFitting,
    'Data Devices': BuiltInCategory.OST_DataDevices,
    'Fire Alarm Devices': BuiltInCategory.OST_FireAlarmDevices,
    'Lighting Devices': BuiltInCategory.OST_LightingDevices,
    'Nurse Call Devices': BuiltInCategory.OST_NurseCallDevices,
    'Security Devices': BuiltInCategory.OST_SecurityDevices,
    'Telephone Devices': BuiltInCategory.OST_TelephoneDevices,

    'Walls': BuiltInCategory.OST_Walls,
    'Ceilings': BuiltInCategory.OST_Ceilings,

    'Columns': BuiltInCategory.OST_StructuralColumns,
    'Framing': BuiltInCategory.OST_StructuralFraming,
    'Floors': BuiltInCategory.OST_Floors,
}

# Main Code
# =================================================

# Show selection dialog
selected_names = forms.SelectFromList.show(
    sorted(category_options.keys()),
    title='Select Categories to Isolate',
    multiselect=True,
    button_name='Isolate Selected'
)

if not selected_names:
    sys.exit(0)

# Collect element ids from selected categories
ids = List[ElementId]()
for name in selected_names:
    bic = category_options.get(name)
    if not bic:
        continue
    collector = FilteredElementCollector(doc, active_view.Id).OfCategory(
        bic).WhereElementIsNotElementType()
    for el in collector:
        ids.Add(el.Id)

# If nothing collected, exit silently
if ids.Count == 0:
    sys.exit(0)

# Apply temporary isolation within a transaction
with revit.Transaction('Isolate Selected Categories'):
    active_view.IsolateElementsTemporary(ids)
