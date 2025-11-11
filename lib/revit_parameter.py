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
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')   # optional if you use RevitServices, adjust as needed

from Autodesk.Revit.DB import StorageType, ExternalDefinitionCreationOptions, BuiltInCategory, ElementId
from Autodesk.Revit.DB import CategorySet  # only if you need it explicitly

from pyrevit import revit, forms

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
            raise LookupError("Parameter '{}' not found on element.".format(param_name))

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
            raise TypeError("Unsupported type for parameter '{}'".format(param_name))


    def create_parameter(self,
                         name,                                                      # Name of Paramter
                         param_type         = None,                   # Type: Text, Number, YesNo, Length, etc.
                         group_name         = "MyGroup",                            # Name of the group
                         categories_to_bind = None,                                 # OST_Walls, OST_Doors, etc. defaluts to OST_DuctCurves
                         is_instance        = True,                                 # True = Instance, False = Type; parameter wise
                         param_group        = None):    # Group in properties palette, DATA, GEOMETRY, etc.
        
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
            catset.Insert(self.doc.Settings.Categories.get_Item(BuiltInCategory.OST_DuctCurves))

        # 5) Choose binding type
        binding = (self.app.Create.NewInstanceBinding(catset)
                   if is_instance else
                   self.app.Create.NewTypeBinding(catset))

        # 6) Insert or update binding
        binding_map = self.doc.ParameterBindings
        if not binding_map.Insert(definition, binding):
            binding_map.ReInsert(definition, binding)

        return definition
