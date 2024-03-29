# -*- coding: utf-8 -*-
__title__ = "Maulweite"
__author__ = "olga.poletkina@hpp.com"
__doc__ = """
Author: olga.poletkina@hpp.com
Date: 18.07.2023
___________________________________________________________
Description:
The script retrieves the host wall width and assigns 
it to the door parameter.
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

from Snippets._functions import unit_converter

# uiapp = __revit__
doc = __revit__.ActiveUIDocument.Document
# uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

doors = [door for door in FEC(doc).OfCategory(
    DB.BuiltInCategory.OST_Doors).WhereElementIsNotElementType().ToElements() if door.Parameter[
    DB.BuiltInParameter.PHASE_DEMOLISHED].AsDouble() == 0]

with DB.Transaction(doc, 'Wall thickness parameter application') as t:
    t.Start()
    width = []
    for door in doors:
        thickness_parameter = door.LookupParameter('H_TÜ_ZA_Maulweite')
        if thickness_parameter == None:
            print('Please apply "H_TÜ_ZA_Maulweite" parameter!')
            break
        elif thickness_parameter != None:
            thickness_parameter.Set(str(unit_converter(doc, door.Host.Width)))
            width.append(unit_converter(doc, door.Host.Width))
    t.Commit()

print("""The 'H_TÜ_ZA_Maulweite' parameter is applied.""")

