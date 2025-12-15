# -*- coding: utf-8 -*-
############################################################################
# Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.
#
# This code and associated documentation files may not be copied, modified,
# distributed, or used in any form without the prior written permission of
# the copyright holder.
############################################################################

# Imports
# ==========================================================================
from pyrevit import revit, forms
from Autodesk.Revit.DB import (
    CategorySet, StorageType, ExternalDefinitionCreationOptions,
    BuiltInCategory, ElementId, ParameterTypeId
)
import clr
clr.AddReference('RevitAPI')
# optional if you use RevitServices, adjust as needed
clr.AddReference('RevitServices')


# Variables
# ==========================================================================
app = __revit__.Application
doc = revit.doc

# Class
# ==========================================================================


class RevitParameter:
    def __init__(self, doc, app):
        self.doc = doc
        self.app = app

    def get_parameter_value(self, element, param_name):
        """Retrieve the value of a parameter by name from a given element."""
        param = element.LookupParameter(param_name)
        if not param:
            raise LookupError(
                "Parameter '{}' not found on element.".format(param_name))

        st = param.StorageType
        if st == StorageType.String:
            return param.AsString()
        elif st == StorageType.Integer:
            return param.AsInteger()
        elif st == StorageType.Double:
            return param.AsDouble()
        elif st == StorageType.ElementId:
            return param.AsElementId()
        else:
            return None

    def set_parameter_value(self, element, param_name, value):
        param = element.LookupParameter(param_name)
        if not param:
            raise LookupError("Parameter '{}' not found".format(param_name))

        st = param.StorageType
        if st == StorageType.String:
            param.Set(str(value))
        elif st == StorageType.Integer:
            param.Set(int(value))
        elif st == StorageType.Double:
            param.Set(float(value))
        elif st == StorageType.ElementId and isinstance(value, ElementId):
            param.Set(value)
        else:
            raise TypeError(
                "Unsupported type for parameter '{}'".format(param_name))

    def create_parameter(self,
                         name,  # Name of Paramter
                         # Type: Text, Number, YesNo, Length, etc.
                         param_type=None,
                         group_name="MyGroup",                            # Name of the group
                         categories_to_bind=None,  # OST_Walls, OST_Doors, etc. defaluts to OST_DuctCurves
                         # True = Instance, False = Type; parameter wise
                         is_instance=True,
                         param_group=None):    # Group in properties palette, DATA, GEOMETRY, etc.

        # Convert string param_type to ForgeTypeId if needed
        if isinstance(param_type, str):
            # Revit 2026 uses ForgeTypeId for parameter types
            try:
                # ParameterTypeId is a sealed class with static properties
                # that return ForgeTypeId objects
                param_lower = param_type.lower()

                if param_lower == 'text':
                    param_type = ParameterTypeId.Text
                elif param_lower == 'length':
                    param_type = ParameterTypeId.Length
                elif param_lower == 'number':
                    param_type = ParameterTypeId.Number
                else:
                    # Default to text for unknown types
                    param_type = ParameterTypeId.Text
            except Exception as e:
                # If that fails, try ForgeTypeId.Parse as fallback
                try:
                    from Autodesk.Revit.DB import ForgeTypeId
                    schema_map = {
                        'text': 'autodesk.revit.db.parameters:type-text',
                        'length': 'autodesk.revit.db.parameters:type-length',
                        'number': 'autodesk.revit.db.parameters:type-number',
                    }
                    schema = schema_map.get(
                        param_type.lower(), 'autodesk.revit.db.parameters:type-text')
                    param_type = ForgeTypeId.Parse(schema)
                except Exception as parse_error:
                    # Last resort - pass the string and let it fail with clear error
                    pass

        # 1) Open shared parameter file
        spfile = self.app.OpenSharedParameterFile()
        if not spfile:
            forms.alert("No shared parameter file is set in Revit.")
            return None

        # 2) Get or create group
        group = spfile.Groups.get_Item(group_name)
        if not group:
            group = spfile.Groups.Create(group_name)

        # 3) Find or create definition
        definition = None
        for defn in group.Definitions:
            if defn.Name == name:
                definition = defn
                break
        if not definition:
            options = ExternalDefinitionCreationOptions(name, param_type)
            definition = group.Definitions.Create(options)

        # 4) Build category set
        catset = self.app.Create.NewCategorySet()
        if categories_to_bind:
            for bic in categories_to_bind:
                catset.Insert(self.doc.Settings.Categories.get_Item(bic))
        else:
            # Default example: ducts
            catset.Insert(self.doc.Settings.Categories.get_Item(
                BuiltInCategory.OST_DuctCurves))

        # 5) Choose binding type
        binding = (self.app.Create.NewInstanceBinding(catset)
                   if is_instance else
                   self.app.Create.NewTypeBinding(catset))

        # 6) Insert or update binding
        binding_map = self.doc.ParameterBindings
        if not binding_map.Insert(definition, binding):
            binding_map.ReInsert(definition, binding)

        return definition
