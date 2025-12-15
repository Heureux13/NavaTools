# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

# Imports
# ==================================================
import Autodesk.Revit.DB as DB
from revit_output import print_disclaimer
from pyrevit import revit, script
from Autodesk.Revit.DB import BuiltInCategory, StorageType

# Button info
# ===================================================
__title__ = "Create Project Parameters"
__doc__ = """
Creates project parameters for duct documentation.
"""

# Variables
# ==================================================
app = __revit__.Application
doc = revit.doc
output = script.get_output()

# Parameter definitions
# Format: (name, storage_type, revit_categories)

mep_fabrication_hangers = [
    ('_ha_weight_support', StorageType.Double,
     [BuiltInCategory.OST_FabricationHangers]),
]

mep_fabrication_ductwork = [
    ('_du_fabrication_service', StorageType.String,
     [BuiltInCategory.OST_DuctCurves]),
    ('_du_insulation_specification', StorageType.String,
     [BuiltInCategory.OST_DuctCurves]),
    ('_du_offset_center_h', StorageType.Double,
     [BuiltInCategory.OST_DuctCurves]),
    ('_du_offset_center_v', StorageType.Double,
     [BuiltInCategory.OST_DuctCurves]),
    ('_du_offset_down', StorageType.Double,
     [BuiltInCategory.OST_DuctCurves]),
    ('_du_offset_left', StorageType.Double,
     [BuiltInCategory.OST_DuctCurves]),
    ('_du_offset_right', StorageType.Double,
     [BuiltInCategory.OST_DuctCurves]),
    ('_du_offset_up', StorageType.Double,
     [BuiltInCategory.OST_DuctCurves]),
    ('_du_size', StorageType.String,
     [BuiltInCategory.OST_DuctCurves]),
    ('_du_tag_offset', StorageType.Double,
     [BuiltInCategory.OST_DuctCurves]),
    ('_du_weight_run', StorageType.Double,
     [BuiltInCategory.OST_DuctCurves]),
]

mechanical_equipment = [
    ('_eq_damper', StorageType.String,
     [BuiltInCategory.OST_MechanicalEquipment]),
    ('_eq_handing_control', StorageType.String,
     [BuiltInCategory.OST_MechanicalEquipment]),
    ('_eq_handing_piping', StorageType.String,
     [BuiltInCategory.OST_MechanicalEquipment]),
    ('_eq_label', StorageType.String,
     [BuiltInCategory.OST_MechanicalEquipment]),
    ('_eq_make', StorageType.String,
     [BuiltInCategory.OST_MechanicalEquipment]),
    ('_eq_model', StorageType.String,
     [BuiltInCategory.OST_MechanicalEquipment]),
    ('_eq_mount', StorageType.String,
     [BuiltInCategory.OST_MechanicalEquipment]),
    ('_eq_opening_inlet_ea', StorageType.String,
     [BuiltInCategory.OST_MechanicalEquipment]),
    ('_eq_opening_inlet_oa', StorageType.String,
     [BuiltInCategory.OST_MechanicalEquipment]),
    ('_eq_opening_inlet_ra', StorageType.String,
     [BuiltInCategory.OST_MechanicalEquipment]),
    ('_eq_opening_inlet_sa', StorageType.String,
     [BuiltInCategory.OST_MechanicalEquipment]),
    ('_eq_opening_outlet_ea', StorageType.String,
     [BuiltInCategory.OST_MechanicalEquipment]),
    ('_eq_opening_outlet_oa', StorageType.String,
     [BuiltInCategory.OST_MechanicalEquipment]),
    ('_eq_opening_outlet_ra', StorageType.String,
     [BuiltInCategory.OST_MechanicalEquipment]),
    ('_eq_opening_outlet_sa', StorageType.String,
     [BuiltInCategory.OST_MechanicalEquipment]),
    ('_eq_size', StorageType.String,
     [BuiltInCategory.OST_MechanicalEquipment]),
    ('_eq_type', StorageType.String,
     [BuiltInCategory.OST_MechanicalEquipment]),
    ('_eq_volt', StorageType.String,
     [BuiltInCategory.OST_MechanicalEquipment]),
]

# Combine all parameters in alphabetical order by category
all_parameters = mep_fabrication_hangers + \
    mep_fabrication_ductwork + mechanical_equipment

# Combine all parameters in alphabetical order by category
all_parameters = mep_fabrication_hangers + \
    mep_fabrication_ductwork + mechanical_equipment

# Main Code - Create Project Parameters
# =================================================

created = []
failed = []

with revit.Transaction("Create Project Parameters"):
    for param_name, storage_type, categories in all_parameters:
        try:
            param_bindings = doc.ParameterBindings

            # Build category set
            cat_set = app.Create.NewCategorySet()
            for bic in categories:
                cat = doc.Settings.Categories.get_Item(bic)
                cat_set.Insert(cat)

            # Create instance binding for these parameters
            binding = app.Create.NewInstanceBinding(cat_set)

            # Check if parameter already exists by iterating through definitions
            param_exists = False
            for param_id in param_bindings:
                # param_id is actually the ExternalDefinition object
                if param_id.Name == param_name:
                    param_exists = True
                    # Update binding if exists
                    param_bindings.ReInsert(param_id, binding)
                    break

            if not param_exists:
                # Create new external definition
                external_def = param_bindings.Create(
                    param_name,
                    storage_type
                )

                # Insert the definition with its binding
                param_bindings.Insert(external_def, binding)

            created.append(param_name)
            output.print_md("✓ Created: {}".format(param_name))
        except Exception as e:
            failed.append((param_name, str(e)))
            output.print_md("✗ Failed: {} - {}".format(param_name, str(e)))

output.print_md("---")
output.print_md("## Summary")
output.print_md("Created: {} | Failed: {}".format(len(created), len(failed)))
print_disclaimer(output)
