# -*- coding: utf-8 -*-
__title__ = "Floor Join"
__author__ = "olga.poletkina@hpp.com"
__doc__ = """
Author: olga.poletkina@hpp.com
Date: 18.07.2023
___________________________________________________________
Description:
This script joins specific categories (stairs, columns, 
foundations, framing, and walls) with finishing floors 
that are visible in the active view. It filters the visible 
elements, identifies intersections, and performs the join 
operation. The output indicates if any joins were made or 
if there were no intersections in the visible elements.
___________________________________________________________
How-to:
Press the button.
___________________________________________________________
Prerequisite:
Please limit the visibility or selection of elements to a 
specific area or subset before running the script. 
Running the script on the entire file may 
cause performance issues!
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

# uiapp = __revit__
doc = __revit__.ActiveUIDocument.Document
# uidoc = __revit__.ActiveUIDocument
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

def flatten(element, flat_list=None):
    if flat_list is None:
        flat_list = []
    if hasattr(element, "__iter__"):
        for item in element:
            flatten(item, flat_list)
    else:
        flat_list.append(element)
    return flat_list

categories_filter = List[DB.BuiltInCategory]()

# add categoties to the list
categories_filter.Add(DB.BuiltInCategory.OST_Stairs)
categories_filter.Add(DB.BuiltInCategory.OST_StructuralColumns)
categories_filter.Add(DB.BuiltInCategory.OST_StructuralFoundation)
categories_filter.Add(DB.BuiltInCategory.OST_StructuralFraming)
categories_filter.Add(DB.BuiltInCategory.OST_Walls)

# create multi category filter
multi_category_filter = DB.ElementMulticategoryFilter(categories_filter)

# filter for floors
# string contains rule
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

# final LogicalOr filter
or_filter = DB.LogicalOrFilter(
    to_list(
        [
            multi_category_filter,
            and_rule_floors
        ],
        DB.ElementFilter
    )
)

elements = FEC(doc, doc.ActiveView.Id).WherePasses(multi_category_filter).WhereElementIsNotElementType().ToElements()
floors = FEC(doc, doc.ActiveView.Id).WherePasses(and_rule_floors).WhereElementIsNotElementType().ToElements()

cancelled = False
intersected = []

# get intersected elements and initiate progress bar
with ProgressBar(cancellable=True) as pb:
    elements_from_filter = []
    intersected_list = []
    for floor, counter in zip(floors, range(len(floors))):
        intersect_filter = DB.ElementIntersectsElementFilter(floor)
        elements_from_filter = FEC(doc, doc.ActiveView.Id).WherePasses(multi_category_filter).WherePasses(intersect_filter)
        intersected_list.append(elements_from_filter)
        intersected_Ids = []
        for el in flatten(intersected_list):
            intersected_Ids.append(doc.GetElement(el.Id))
        if intersected_Ids:
            intersected.append([floor, intersected_Ids])
        if pb.cancelled:
            cancelled = True
            break
        else:
            pb.update_progress(counter, len(elements))

if cancelled:
    print('Operation is cancelled!')

# join elements
with DB.Transaction(doc, 'Join elements') as t:
    t.Start()
    for element in intersected:
        try:
            for el in element[1]:
                DB.JoinGeometryUtils.JoinGeometry(
                    doc,
                    element[0],
                    el
                )
        except:
            pass
    t.Commit()

# switch joining order
with DB.Transaction(doc, 'Switch join order for finished floor') as t:
    t.Start()
    for element in intersected:
        joined_elements = DB.JoinGeometryUtils.GetJoinedElements(doc, element[0])
        for el_id in joined_elements:
            el = doc.GetElement(el_id)
            if DB.JoinGeometryUtils.IsCuttingElementInJoin(doc, element[0], el):
                try:
                    DB.JoinGeometryUtils.SwitchJoinOrder(
                        doc,
                        el,
                        element[0]
                )
                except:
                    pass
    t.Commit()

if not cancelled:
    if len(intersected) > 0:
        print('Elements were joined!')
    else:
        print('No intersections were detected!')

