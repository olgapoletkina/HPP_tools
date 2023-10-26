# -*- coding: utf-8 -*-
__title__ = "Join All"
__author__ = "olga.poletkina@hpp.com"
__doc__ = """
Author: olga.poletkina@hpp.com
Date: 18.07.2023
___________________________________________________________
Description:
This script retrieves elements of specific categories 
such as stairs, structural columns, structural foundation, 
structural framing, walls, and floors from the active view. 
It then checks for intersecting or adjacent elements and 
joins them using Revit's JoinGeometryUtils. 
The script provides an output message indicating 
that the elements have been successfully joined.
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

uiapp = __revit__
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

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

def flatten(element, flat_list=None):
    if flat_list is None:
        flat_list = []
    if hasattr(element, "__iter__"):
        for item in element:
            flatten(item, flat_list)
    else:
        flat_list.append(element)
    return flat_list

g_options = DB.Options()

categories_filter = List[DB.BuiltInCategory]()
# add categoties to the list
categories_filter.Add(DB.BuiltInCategory.OST_Stairs)
categories_filter.Add(DB.BuiltInCategory.OST_StructuralColumns)
categories_filter.Add(DB.BuiltInCategory.OST_StructuralFoundation)
categories_filter.Add(DB.BuiltInCategory.OST_StructuralFraming)
categories_filter.Add(DB.BuiltInCategory.OST_Walls)
categories_filter.Add(DB.BuiltInCategory.OST_Floors)
categories_filter.Add(DB.BuiltInCategory.OST_Ceilings)

# create multi category filter
multi_category_filter = DB.ElementMulticategoryFilter(categories_filter)

elements = FEC(doc, doc.ActiveView.Id).WherePasses(multi_category_filter).WhereElementIsNotElementType().ToElements()

cancelled = False
# create a dictionary to store intersecting elements
intersections = {}
with ProgressBar(cancellable=True) as pb:
    # iterate through the elements and check for intersecting bounding boxes
    for [i, el1], counter in zip(enumerate(elements), range(0, len(elements))):
        # get the element's bounding box
        bbox1 = el1.get_BoundingBox(None)
        for j in range(i + 1, len(elements)):
            el2 = elements[j]
            # get the other element's bounding box
            bbox2 = el2.get_BoundingBox(None)
            # check if the two bounding boxes intersect
            if bbox1 != None:
                outline1 = DB.Outline(bbox1.Min, bbox1.Max)
            if bbox2 != None:
                outline2 = DB.Outline(bbox2.Min, bbox2.Max)
            if outline1.Intersects(outline2, 10):
                if el1.Id not in intersections:
                    intersections[el1.Id] = []
                intersections[el1.Id].append(el2.Id)
                if el2.Id not in intersections:
                    intersections[el2.Id] = []
                intersections[el2.Id].append(el1.Id)
            if pb.cancelled:
                cancelled = True
                break
            else:
                pb.update_progress(counter, len(elements))

if cancelled:
    print('Operation is cancelled!')

# list of elements to join
inter = []
for el_id, intersecting_ids in intersections.items():
    inter.append([el_id, intersecting_ids])

with DB.Transaction(doc, 'Join Elements') as t:
    t.Start()
    for i, el in enumerate(inter):
        for j, el_l in enumerate(inter[i][1]):
            try:
                DB.JoinGeometryUtils.JoinGeometry(
                    doc,
                    doc.GetElement(el[0]),
                    doc.GetElement(el_l)
                )
            except:
                pass
    t.Commit()

if not cancelled:
    print('Elements are joined!')