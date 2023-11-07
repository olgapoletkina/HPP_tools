# -*- coding: utf-8 -*-

import sys
import clr
clr.AddReference('ProtoGeometry')
clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
clr.AddReference('RevitNodes')
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)
clr.AddReference('RevitAPI')
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector as FEC
from System.Collections.Generic import List

uiapp = __revit__
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

def flatten(element, flat_list=None):

    '''gets the list with complex structure, 
    returns flattened list of elements'''

    if flat_list is None:
        flat_list = []
    if hasattr(element, "__iter__"):
        for item in element:
            flatten(item, flat_list)
    else:
        flat_list.append(element)
    return flat_list

def to_list(element, list_type=None):
    if not hasattr(element, '__iter__') or isinstance(element, dict):
        element = [element]
    if list_type is not None:
        if isinstance(element, List[list_type]):
            return element
        if all(isinstance(item, list_type) for item in element):
            typed_list = List[list_type]()
            for item in element:
                typed_list.Add(item)
            return typed_list
    return element

def create_parameter_binding(doc, categories, is_type_binding=False):
    app = doc.Application
    category_set = app.Create.NewCategorySet()
    for category in to_list(categories):
        if isinstance(category, DB.BuiltInCategory):
            category = DB.Category.GetCategory(doc, category)
        category_set.Insert(category)
    if is_type_binding:
        return app.Create.NewTypeBinding(category_set)
    return app.Create.NewInstanceBinding(category_set)

def create_project_parameter(doc, external_definition, binding, p_group = DB.BuiltInParameterGroup.INVALID):
    if doc.ParameterBindings.Insert(external_definition, binding, p_group):
        iterator = doc.ParameterBindings.ForwardIterator()
        while iterator.MoveNext():
            internal_definition = iterator.Key
            parameter_element = doc.GetElement(internal_definition.Id)
            if isinstance(parameter_element, DB.SharedParameterElement) and parameter_element.GuidValue == external_definition.GUID:
                return internal_definition

def get_external_definition(app, group_name, definition_name):
    return app.OpenSharedParameterFile() \
        .Groups[group_name] \
        .Definitions[definition_name]