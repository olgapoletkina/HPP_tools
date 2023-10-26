# -*- coding: utf-8 -*-
__title__ = "Room Number"
__author__ = "olga.poletkina@hpp.com"
__doc__ = """
Author: olga.poletkina@hpp.com
Date: 18.07.2023
___________________________________________________________
Description:
This script generates and assigns room numbers based on 
the room name and existing number parameter. 
It also checks for rooms that do not have the 
"H_RA_Raumnummer" parameter or have an empty value, 
and prompts to review the "Raumliste".
___________________________________________________________
How-to:
Press the button.
___________________________________________________________
Prerequisite:
Please ensure that the 'Number' and 'Name' parameters are 
applied before executing the script!!!
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

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

rooms = [room for room in FEC(doc).OfCategory(
        BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements() if room.Parameter[
        DB.BuiltInParameter.ROOM_PHASE].AsValueString() == 'Neu' and room.Parameter[
        DB.BuiltInParameter.ROOM_AREA].AsDouble() != 0]

# for room in rooms:
#     if room.LookupParameter('H_RA_Raumnummer') == None:
#         raise Exception ('Please apply parameter "H_RA_Raumnummer" to rooms!')

for room in rooms:
    if room.LookupParameter('H_RA_Raumnummer') == None:
        print('Please apply parameter "H_RA_Raumnummer" to room!')
        break

numbers_generated = []
temp_numbers = []
temp_room = []
with DB.Transaction(doc, 'Assign Room Number') as t:
    t.Start()
    for room in rooms:
        number_param = room.Parameter[DB.BuiltInParameter.ROOM_NUMBER].AsString()
        name_param = room.Parameter[DB.BuiltInParameter.ROOM_NAME].AsString()
        room_number = number_param + '.' + name_param[:1].upper()
        if room_number in numbers_generated:
            temp_numbers.append(room_number)
            temp_room.append(room)
        else:
            numbers_generated.append(room_number)
            paumnummer_param = room.LookupParameter('H_RA_Raumnummer')
            paumnummer_param.Set(room_number)
    for room, [i, value] in zip(temp_room, enumerate(temp_numbers)):
        totalcount = temp_numbers.count(value)
        count = temp_numbers[:i].count(value)
        room_number = value + str(count + 1) if totalcount > 0 else value
        paumnummer_param = room.LookupParameter('H_RA_Raumnummer')
        paumnummer_param.Set(room_number)
        numbers_generated.append(room_number)
    t.Commit()
    
print('The following room numbers were generated and applied:')
for number in numbers_generated:
    print(number)

all_rooms = [room for room in FEC(doc).OfCategory(
        BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements() if room.Parameter[
        DB.BuiltInParameter.ROOM_PHASE].AsValueString() == 'Neu']
if len(all_rooms) < 0:
    print('***')
    print('The following rooms got no "H_RA_Raumnummer" parameter:')
    for room in all_rooms:
        if room.LookupParameter(
            'H_RA_Raumnummer').AsString() == None or room.LookupParameter(
            'H_RA_Raumnummer').AsString() == '' or room.Parameter[
            DB.BuiltInParameter.ROOM_AREA].AsDouble() == 0:
            print(room.Id)
    print('Check "Raumliste"!')

