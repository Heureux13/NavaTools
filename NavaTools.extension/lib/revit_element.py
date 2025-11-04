# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import *
from pyrevit import revit, forms, DB
from Autodesk.Revit.UI import UIDocument
from Autodesk.Revit.ApplicationServices import Application
import re

# Variables
app   = __revit__.Application #type: Application
uidoc = __revit__.ActiveUIDocument #type: UIDocument
doc   = revit.doc #type: Document
view  = revit.active_view

class RevitElement:
    def __init__ (self, doc, view, element):
        self.doc = doc
        self.view = view
        self.element = element

    @property
    def id(self):
        return self.element.Id
    
    @property
    def category(self):
        return self.element.Category.Name if self.element and self.element.Category else None
    
    def get_param(self, param_name):
        param = self.element.LookupParameter(param_name)
        if param:
            if param.StorageType == DB.StorageType.String:
                return param.AsString()
            elif param.StorageType == DB.StorageType.Double:
                return param.AsDouble()
            elif param.StorageType == DB.StorageType.Integer:
                return param.AsInteger()
            elif param.StorageType == DB.StorageType.ElementId:
                return param.AsElementId()
        return None
    
    def set_param(self, param_name, value):
        param = self.element.LookupParameter(param_name)
        if not param:
            print("Parameter '{}' not found on element {}".format(param_name, self.id))
            return
        if param.IsReadOnly:
            print("Parameter '{}' is read-only on element {}".format(param_name, self.id))
            return
        
        if isinstance(value, str):
            param.Set(value)
        elif isinstance(value, int):
            param.Set(int(value))
        elif isinstance(value, float):
            param.Set(float(value))
        elif isinstance(value, DB.ElementId):
            param.Set(value)
        else:
            print("Unsupported value type for parameter '{}' on element {}".format(param_name, self.id))