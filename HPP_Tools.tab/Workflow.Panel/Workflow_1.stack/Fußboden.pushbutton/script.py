# -*- coding: utf-8 -*-
__title__ = "Fußboden"
__author__ = "olga.poletkina@hpp.com"
__doc__ = """
Author: olga.poletkina@hpp.com
Date: 18.07.2023
___________________________________________________________
Description:
This tool applies the "H_TÜ_Fußbodenaufbau" parameter to 
the door. It does this by first setting all door 
"Sill height" parameters to zero. Next, the tool creates a 
bounding box for floors that contain strings "GFB", "GDA", 
or "DAD" in their names. It then creates bounding boxes 
for the doors. If the door bounding box intersects with 
the floor bounding box, the floor thickness parameter 
is used to apply it to the "H_TÜ_Fußbodenaufbau" parameter.
___________________________________________________________
How-to:
Press the button.
___________________________________________________________
Prerequisite:
The corresponding parameter should be applied!
___________________________________________________________
"""
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

import pyrevit
from pyrevit.forms import ProgressBar

# doc = DocumentManager.Instance.CurrentDBDocument
# uiapp = DocumentManager.Instance.CurrentUIApplication
# app = uiapp.Application
# uidoc = uiapp.ActiveUIDocument

# uiapp = __revit__
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

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

def flatten(element, flat_list=None):
    if flat_list is None:
        flat_list = []
    if hasattr(element, "__iter__"):
        for item in element:
            flatten(item, flat_list)
    else:
        flat_list.append(element)
    return flat_list

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

doors = [door for door in FEC(doc).OfCategory(
    DB.BuiltInCategory.OST_Doors).WhereElementIsNotElementType().ToElements() if door.Parameter[
    DB.BuiltInParameter.PHASE_CREATED].AsValueString() == 'Neu']

# rules for floors
gfb_contains_rule = DB.ParameterFilterRuleFactory.CreateContainsRule(
    DB.ElementId(DB.BuiltInParameter.ELEM_TYPE_PARAM),
    'GFB',
    True
)
gda_contains_rule = DB.ParameterFilterRuleFactory.CreateContainsRule(
    DB.ElementId(DB.BuiltInParameter.ELEM_TYPE_PARAM),
    'GDA',
    True
)
dad_contains_rule = DB.ParameterFilterRuleFactory.CreateContainsRule(
    DB.ElementId(DB.BuiltInParameter.ELEM_TYPE_PARAM),
    'DAD',
    True
)
floor_category_rule = DB.FilterCategoryRule(
    to_list(
        DB.ElementId(DB.BuiltInCategory.OST_Floors),
        DB.ElementId
    )
)
# parameter or filter
parameter_contains_filter = DB.LogicalOrFilter(
    to_list(
        [
            DB.ElementParameterFilter(gfb_contains_rule),
            DB.ElementParameterFilter(gda_contains_rule),
            DB.ElementParameterFilter(dad_contains_rule)
        ],
        DB.ElementFilter
    )
)
# floors filter
and_rule_floors = DB.LogicalAndFilter(
    to_list(
        [
            parameter_contains_filter,
            DB.ElementParameterFilter(floor_category_rule)
        ],
        DB.ElementFilter
    )
)

floors = FEC(doc).WherePasses(and_rule_floors).WhereElementIsNotElementType().ToElements()

no_param = []
with DB.Transaction(doc, 'Sill Height application') as t:
    t.Start()
    for door in doors:
        sill_height = door.Parameter[DB.BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM]
        sill_height.Set(0)
        if door.LookupParameter('H_TÜ_Fußbodenaufbau') != None:
            door_fuss_param = door.LookupParameter('H_TÜ_Fußbodenaufbau')
            door_fuss_param.Set(0)
        else:
            if door.Symbol.Id not in no_param:
                no_param.append(door.Symbol.Id)
    t.Commit()

intersecting = []
not_intersecting = []
cancelled = False

# iterate through the elements and check for intersecting bounding boxes
with DB.Transaction(doc, 'Fußboden application') as t:
    t.Start()
    with ProgressBar(cancellable=True) as pb:
        for door, counter in zip(doors, range(len(doors))):
            door_fuss_param = door.LookupParameter('H_TÜ_Fußbodenaufbau')
            bbox_door = door.get_BoundingBox(None)
            for floor in floors:
                bbox_floor = floor.get_BoundingBox(None)
                if floor.Parameter[DB.BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM] != None:
                    floor_param_double = floor.Parameter[DB.BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM].AsDouble()
                    if bbox_door != None and bbox_floor != None: 
                        # outline_door = DB.Outline(bbox_door.Min - DB.XYZ(0, 0, unit_conventer(doc, 0.5, to_internal=True)), bbox_door.Max)
                        outline_door = DB.Outline(bbox_door.Min, bbox_door.Max)
                        outline_floor = DB.Outline(bbox_floor.Min, bbox_floor.Max)
                        if outline_door.Intersects(outline_floor, 0) == True:
                            # print(outline_door.Intersects(outline_floor, 0))
                            if door_fuss_param != None and floor.Parameter[DB.BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM] != None:
                                door_fuss_param.Set(floor_param_double)
                            elif door.Id not in intersecting:
                                intersecting.append(door.Id)
                        elif outline_door.Intersects(outline_floor, 0) == False:
                            pass
            
            pb.update_progress(counter, len(doors))
            
            if pb.cancelled:
                cancelled = True
                break

    if cancelled:
        print('Operation is cancelled!')

    t.Commit()

if not cancelled:
    if len(no_param) > 0:
        with DB.Transaction(doc, 'Create list "Fußbodenaufbau Check"') as t:
            t.Start()
            if len([item for item in FEC(doc).OfClass(DB.ViewSchedule).ToElements() if item.Name == 'Fußbodenaufbau Check']) > 0:
                print('Existing schedule "Fußbodenaufbau Check" is modified.')
                print('*****')
            else:
                door_schedule = DB.ViewSchedule.CreateSchedule(
                                doc,
                                DB.ElementId(DB.BuiltInCategory.OST_Doors)
                            )
                door_schedule.Name = 'Fußbodenaufbau Check'
            # door_schedule = [item for item in FEC(doc).OfClass(DB.ViewSchedule).ToElements() if item.Name == 'Fußbodenaufbau Check'][0]
                s_definition = door_schedule.Definition
                # add schedule fields
                family_field = s_definition.AddField(
                DB.ScheduleFieldType.Instance,
                    DB.ElementId(
                        DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM
                    )
                )
                fussboden_param_id = []
                for door in doors:
                    if door.LookupParameter('H_TÜ_Fußbodenaufbau') != None:
                        fussboden_param_id.append(door.LookupParameter('H_TÜ_Fußbodenaufbau').Id)
                    elif door.LookupParameter('H_TÜ_Fußbodenaufbau') == None:
                        print('Please, apply "H_TÜ_Fußbodenaufbau" parameter')
                        break
                    else:
                        pass
                try:
                    fussboden_field = s_definition.AddField(
                    DB.ScheduleFieldType.Instance,
                        fussboden_param_id[0]
                    )
                    # add schedule filter
                    # fussboden_filter = s_definition.AddFilter(
                    #     DB.ScheduleFilter(
                    #         fussboden_field.FieldId, 
                    #         DB.ScheduleFilterType.HasNoValue
                    #     )
                    # )
                except:
                    pass
            t.Commit()
    else:
        print('Parameter Fußbodenaufbau is applied.')

    if len(no_param)>0:
        print('{} door type(s) doesn`t contain a proper H_TÜ_Fußbodenaufbau parameter or placed incorrectly:'.format(len(no_param)))
        for id in no_param:
            print(doc.GetElement(id).Family.Name)
    else:
        pass
    print('*****')
    print('Check the door list in schedule "Fußbodenaufbau Check".')
