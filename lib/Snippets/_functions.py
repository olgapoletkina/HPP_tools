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

# working with units

def unit_conventer(
        doc,
        value,
        to_internal=False,
        unit_type=DB.SpecTypeId.Length,
        number_of_digits=None):
    display_units = doc.GetUnits().GetFormatOptions(unit_type).GetUnitTypeId()
    method = DB.UnitUtils.ConvertToInternalUnits if to_internal \
        else DB.UnitUtils.ConvertFromInternalUnits
    if number_of_digits is None:
        return method(value, display_units)
    elif number_of_digits > 0:
        return round(method(value, display_units), number_of_digits)
    return int(round(method(value, display_units), number_of_digits))

# working with list structure

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

def group_by_key(elements, key_type='Type'):
    elements = flatten(elements)
    element_groups = {}
    for element in elements:
        if key_type == 'Type':
            key = type(element)
        elif key_type == 'Category':
            for key in DB.BuiltInCategory.GetValues(DB.BuiltInCategory):
                if int(key) == element.Category.Id.IntegerValue:
                    break
        else:
            key = 'Unknown Key'
        if key not in element_groups:
            element_groups[key] = []
        element_groups[key].append(element)
    return element_groups

# working with parameters

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

def get_parameter_value_v2(parameter):
    if isinstance(parameter, DB.Parameter):
        storage_type = parameter.StorageType
        if storage_type:
            exec 'parameter_value = parameter.As{}()'.format(storage_type)
            return parameter_value

# working with bounding boxes

def merge_bounding_boxes(bboxes):
    # type: (list) -> DB.BoundingBoxXYZ
    """Merges multiple bounding boxes"""
    merged_bb = DB.BoundingBoxXYZ()
    merged_bb.Min = DB.XYZ(
        min(bboxes, key=lambda bb: bb.Min.X).Min.X,
        min(bboxes, key=lambda bb: bb.Min.Y).Min.Y,
        min(bboxes, key=lambda bb: bb.Min.Z).Min.Z
    )
    merged_bb.Max = DB.XYZ(
        max(bboxes, key=lambda bb: bb.Max.X).Max.X,
        max(bboxes, key=lambda bb: bb.Max.Y).Max.Y,
        max(bboxes, key=lambda bb: bb.Max.Z).Max.Z
    )
    return merged_bb

def check_intersection(bbox, element):
    """Check if an element's bounding box intersects with the given bounding box."""
    outline = DB.Outline(bbox.Min, bbox.Max)
    filter_intersects = DB.BoundingBoxIntersectsFilter(outline, False)
    if element.get_BoundingBox(None) is not None:
        return filter_intersects.PassesFilter(doc, element.Id)
    return False

# working with solids

def get_all_solids(element, g_options, solids=None):
    '''retrieve all solids from elements'''
    if solids is None:
        solids = []
    if hasattr(element, "Geometry"):
        for item in element.Geometry[g_options]:
            get_all_solids(item, g_options, solids)
    elif isinstance(element, DB.GeometryInstance):
        for item in element.GetInstanceGeometry():
            get_all_solids(item, g_options, solids)
    elif isinstance(element, DB.Solid):
        solids.append(element)
    elif isinstance(element, DB.FamilyInstance):
        for item in element.GetSubComponentIds():
            family_instance = element.Document.GetElement(item)
            get_all_solids(family_instance, g_options, solids)
    return solids

def to_proto_type(elements, of_type=None):
    elements = flatten(elements) if of_type is None \
        else [item for item in flatten(elements) if isinstance(item, of_type)]
    proto_geometry = []
    for element in elements:
        if hasattr(element, 'ToPoint') and element.ToPoint():
            proto_geometry.append(element.ToPoint())
        elif hasattr(element, 'ToProtoType') and element.ToProtoType():
            proto_geometry.append(element.ToProtoType())
    return proto_geometry

# working with elements

def view_exists(doc, view_name):
    views = FEC(doc).OfClass(DB.View).ToElements()
    for view in views:
        if view.Name == view_name:
            return True
    return False

def get_3d_view_type_id(doc):
    collector = FEC(doc).OfClass(DB.ViewFamilyType)
    for view_type in collector:
        if view_type.ViewFamily == DB.ViewFamily.ThreeDimensional:
            return view_type.Id
    return None

def get_room_boundary(doc, item, options):
    e_list = []
    c_list = []
    try:
        for i in item.GetBoundarySegments(options):
            for j in i:
                e_list.append(doc.GetElement(j.ElementId))
                c_list.append(j.Curve.ToProtoType())
    except:
        calculator = DB.SpatialElementGeometryCalculator(doc)
        try:
            results = calculator.CalculateSpatialElementGeometry(item)
            for face in results.GetGeometry().Faces:
                for b_face in results.GetBoundaryFaceInfo(face):
                    e_list.append(doc.GetElement(b_face.SpatialBoundaryElement.HostElementId))
        except:
            pass
    return [e_list, c_list]

