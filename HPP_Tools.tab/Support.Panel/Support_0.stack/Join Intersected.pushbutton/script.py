# -*- coding: utf-8 -*-
__title__ = "Join Intersected"
__author__ = "olga.poletkina@hpp.com"
__doc__ = """
Author: olga.poletkina@hpp.com
Date: 18.07.2023
___________________________________________________________
Description:
This script retrieves certain categories of elements 
(stairs, structural columns, structural foundation, 
structural framing, walls, and floors) in the active view 
and joins them with intersecting elements using Revit's 
JoinGeometryUtils. It provides a message indicating 
if any elements were joined or if no intersections 
were detected.
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

elements = [el for el in FEC(doc, doc.ActiveView.Id).WherePasses(
    multi_category_filter).WhereElementIsNotElementType().ToElements() if el.Parameter[
    DB.BuiltInParameter.PHASE_CREATED].AsValueString() == 'Neu']

cancelled = False
intersected = []

# get intersected elements and initiate progress bar
with ProgressBar(cancellable=True) as pb:
    for element, counter in zip(elements, range(0, len(elements))):
        intersected_elements = set()
        intersect_filter = DB.ElementIntersectsElementFilter(element)
        intersected_elements.update(FEC(doc).WherePasses(multi_category_filter).WherePasses(intersect_filter))
        intersected_Ids = [doc.GetElement(el.Id) for el in intersected_elements]
        
        if intersected_Ids:
            intersected.append([element, intersected_Ids])
        
        pb.update_progress(counter, len(elements))
        
        if pb.cancelled:
            cancelled = True
            break

if cancelled:
    print('Operation is cancelled!')

#  join elements
with DB.Transaction(doc, 'Join elements') as t:
    t.Start()
    for element in intersected:
        for el in element[1]:
            try:
                DB.JoinGeometryUtils.JoinGeometry(
                    doc,
                    element[0],
                    el
                )
            except:
                pass
    t.Commit()
if intersected:
    print('Elements were joined!')
else:
    print('No intersections were detected!')



