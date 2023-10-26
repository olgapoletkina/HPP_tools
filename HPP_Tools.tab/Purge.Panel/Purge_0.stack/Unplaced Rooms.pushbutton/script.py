# -*- coding: utf-8 -*-
__title__ = "Unplaced Rooms"
__author__ = "olga.poletkina@hpp.com"
__doc__ = """
Author: olga.poletkina@hpp.com
Date: 18.07.2023
___________________________________________________________
Description:
This script removes rooms from the Revit project that have 
an area of 0. It collects rooms using the 
BuiltInCategory.OST_Rooms category, identifies the rooms 
with an area of 0, deletes them from the project, 
and provides a list of removed room Ids.
___________________________________________________________
How-to:
Press the button.
___________________________________________________________
Prerequisite:
None
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

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

rooms = FEC(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()

removed_rooms = []
with DB.Transaction(doc, 'Remove Not_Placed and Redundant Rooms') as t:
    t.Start()
    for room in rooms:
        if room.Parameter[DB.BuiltInParameter.ROOM_AREA].AsDouble() == 0:
            removed_rooms.append(room.Id)
            doc.Delete(room.Id)
    t.Commit()

if len(removed_rooms) > 0:
    print('Rooms with following Ids were removed:')
    print(removed_rooms)
else:
    print('No Not_Placed or Redundant rooms were detected!')