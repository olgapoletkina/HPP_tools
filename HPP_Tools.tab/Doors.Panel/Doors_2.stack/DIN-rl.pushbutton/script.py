# -*- coding: utf-8 -*-
__title__ = "DIN-rl"
__author__ = "olga.poletkina@hpp.com"
__doc__ = """
Author: olga.poletkina@hpp.com
Date: 18.07.2023
___________________________________________________________
Description:
This script automates the assignment of the 
'H_TÜ_DIN-rl' parameter to door families based on their 
orientation, ensuring proper DIN Rechts 
or DIN Links alignment.
___________________________________________________________
How-to:
Press the button.
___________________________________________________________
Prerequisite:
The door family should have a left opening orientation!!!
The corresponding parameter should be applied!
___________________________________________________________
"""
import sys
import clr

clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
clr.AddReference('RevitNodes')
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector as FEC

from System import *

# uiapp = __revit__
doc = __revit__.ActiveUIDocument.Document
# uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

doors = FEC(doc).OfCategory(DB.BuiltInCategory.OST_Doors).WhereElementIsNotElementType().ToElements()
one_flip_doors = []
for door in doors:
    if '1FL' in door.Symbol.Family.Name: # check family code in the project and replace with relevant
        if 'TOR' not in door.Symbol.Family.Name:
            if 'SCH' not in door.Symbol.Family.Name:
                one_flip_doors.append(door)
families_in_use_ids = []
for one_flip_door in one_flip_doors:
    if one_flip_door.Symbol.Family.Id not in families_in_use_ids:
        families_in_use_ids.append(one_flip_door.Symbol.Family.Id)
families_in_use = []
for id in families_in_use_ids:
    door_family = doc.GetElement(id)
    families_in_use.append(door_family)

with DB.Transaction(doc, 'Assign Door Opening Parameter') as t:
    t.Start()
    for door in one_flip_doors:
        if door.LookupParameter('H_TÜ_DIN-rl') == None:
            print('Please, apply parameter "H_TÜ_DIN-rl" to the door!')
            break
        else:
            parameter = door.LookupParameter('H_TÜ_DIN-rl')
            facing_flip = door.FacingFlipped
            hand_flip = door.HandFlipped
            if facing_flip == False and hand_flip == True:
                parameter.Set('DIN Rechts')
            elif facing_flip == False and hand_flip == False:
                parameter.Set('DIN Links')
            elif facing_flip == True and hand_flip == False:
                parameter.Set('DIN Rechts')
            elif facing_flip == True and hand_flip == True:
                parameter.Set('DIN Links')
    t.Commit()

if door.LookupParameter('H_TÜ_DIN-rl') != None:
    print("The door 'H_TÜ_DIN-rl' parameter is filled!")


