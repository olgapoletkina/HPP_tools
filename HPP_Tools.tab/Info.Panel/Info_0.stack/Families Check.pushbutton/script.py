# -*- coding: utf-8 -*-
__title__ = "Families Check"
__author__ = "olga.poletkina@hpp.com"
__doc__ = """
Author: olga.poletkina@hpp.com
Date: 18.07.2023
___________________________________________________________
Description:
This script checks if the window and door families 
in a Revit project comply with the HPP naming 
convention and provides a list of families that 
do not meet the convention.
___________________________________________________________
How-to:
Press the button.
___________________________________________________________
Prerequisite:
None.
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

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

windows = FEC(doc).OfCategory(DB.BuiltInCategory.OST_Windows).WhereElementIsNotElementType().ToElements()
doors = FEC(doc).OfCategory(DB.BuiltInCategory.OST_Doors).WhereElementIsNotElementType().ToElements()

check_list = []

for window in windows:
    if 'HA' in window.Symbol.Family.Name or 'HI' in window.Symbol.Family.Name:
        pass
    else:
        if window.Symbol.Family.Name not in check_list:
            check_list.append(window.Symbol.Family.Name)

for door in doors:
    if 'HA' in door.Symbol.Family.Name or 'HI' in door.Symbol.Family.Name:
        pass
    else:
        if door.Symbol.Family.Name not in check_list:
            check_list.append(door.Symbol.Family.Name)

if len(check_list) > 0:
    print('The following families doesn`t comply with the HPP naming convention:')
    for el in check_list:
        print(el)
else:
    print('No conflict with the HPP naming convention was detected.')