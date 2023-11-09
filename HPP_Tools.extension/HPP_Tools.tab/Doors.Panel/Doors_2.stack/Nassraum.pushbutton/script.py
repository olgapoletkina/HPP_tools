# -*- coding: utf-8 -*-
__title__ = "Nassraum"
__author__ = "olga.poletkina@hpp.com"
__doc__ = """
Author: olga.poletkina@hpp.com
Date: 18.07.2023
___________________________________________________________
Description:
This script applies the 'H_TÜ_Nassraum-Feuchtraum' 
parameter to doors, based on whether they belong to 
wet room categories, ensuring accurate
labeling for wet room doors.
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

# doc = DocumentManager.Instance.CurrentDBDocument
# uiapp = DocumentManager.Instance.CurrentUIApplication
# app = uiapp.Application
# uidoc = uiapp.ActiveUIDocument

# uiapp = __revit__
doc = __revit__.ActiveUIDocument.Document
# uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

doors = [door for door in FEC(doc).OfCategory(
    DB.BuiltInCategory.OST_Doors).WhereElementIsNotElementType().ToElements() if door.Parameter[
    DB.BuiltInParameter.PHASE_DEMOLISHED].AsDouble() == 0]

# print(doors)

wet_rooms = [
    'Bad',
    'Toilet',
    'Dusch',
    'WC',
    'Wasch'
]

phase_created = [door.Parameter[DB.BuiltInParameter.PHASE_CREATED].AsElementId() for door in doors]
door_phase = doc.GetElement(phase_created[0])

with DB.Transaction(doc, 'Wet room door parameter application') as t:
    t.Start()
    function = []
    for door in doors:
        rooms = []
        nassraum_parameter = door.LookupParameter('H_TÜ_Nassraum-Feuchtraum')
        if nassraum_parameter is None:
            print('Please apply "H_TÜ_Nassraum-Feuchtraum" parameter!')
            break
        else:
            for string in wet_rooms:
                if door.FromRoom[door_phase] is not None and string in door.FromRoom[
                    door_phase].Parameter[DB.BuiltInParameter.ROOM_NAME].AsString() or door.ToRoom[
                    door_phase] is not None and string in door.ToRoom[
                    door_phase].Parameter[DB.BuiltInParameter.ROOM_NAME].AsString():
                    nassraum_parameter.Set('True')
                    break
                else:
                    nassraum_parameter.Set('False')
                    break
    t.Commit()

if door.LookupParameter('H_TÜ_Nassraum-Feuchtraum') != None:
    print("The door 'H_TÜ_Nassraum-Feuchtraum' parameter is filled!")