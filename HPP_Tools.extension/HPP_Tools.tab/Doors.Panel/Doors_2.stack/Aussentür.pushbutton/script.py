# -*- coding: utf-8 -*-
__title__ = "Aussentür"
__author__ = "olga.poletkina@hpp.com"
__doc__ = """
Author: olga.poletkina@hpp.com
Date: 18.07.2023
___________________________________________________________
Description:
The script checks if the "H_TÜ_Aussentür" parameter is 
present on doors and sets its value based on the 
"FUNCTION_PARAM" value of the door's symbol. 
It also provides feedback regarding the application 
of the "H_TÜ_Aussentür" parameter.
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

# uiapp = __revit__
doc = __revit__.ActiveUIDocument.Document
# uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

doors = [door for door in FEC(doc).OfCategory(
    DB.BuiltInCategory.OST_Doors).WhereElementIsNotElementType().ToElements() if door.Parameter[
    DB.BuiltInParameter.PHASE_DEMOLISHED].AsDouble() == 0]
phase_new = [phase for phase in FEC(doc).OfClass(DB.Phase) if phase.Name == 'Neu' or phase.Name == 'New']
phase_demo = [phase for phase in FEC(doc).OfClass(DB.Phase) if phase.Name == 'Bestand'] #or phase.Name == 'Demolishion' - english name?
# print(phase_new[0])
# for p in phase_new:
#     print(p.Name)

with DB.Transaction(doc, 'Outside or inside door parameter application') as t:
    t.Start()
    function = []
    for door in doors:
        if door.ToRoom[phase_new[0]] != None and door.FromRoom[phase_new[0]] != None:
            # print([door.ToRoom[phase_new[0]].Id, door.FromRoom[phase_new[0]].Id, 'inside'])
            comments_param = door.LookupParameter('Comments')
            comments_param.Set('INSIDE')
        elif door.ToRoom[phase_new[0]] == None and door.FromRoom[phase_new[0]] != None or \
            door.ToRoom[phase_new[0]] != None and door.FromRoom[phase_new[0]] == None:
            # print([door.ToRoom[phase_new[0]].Id, door.FromRoom[phase_new[0]].Id, 'outside'])
            comments_param = door.LookupParameter('Comments')
            comments_param.Set('OUTSIDE')
        # else:
        #     comments_param.Set('OUTSIDE DOOR')
            # print('INSIDE DOOR')
        in_out_parameter = door.LookupParameter('H_TÜ_Aussentür')
        if in_out_parameter == None:
            print('Please apply "H_TÜ_Aussentür" parameter!')
            break
        elif in_out_parameter != None:
            if door.Symbol.Parameter[DB.BuiltInParameter.FUNCTION_PARAM].AsInteger() == 0:
                in_out_parameter.Set(False)
            else:
                in_out_parameter.Set(True)
    t.Commit()

for door in doors:
    if door.LookupParameter('H_TÜ_Aussentür') is not None:
        print('"H_TÜ_Aussentür" parameter is applied!')
        break
