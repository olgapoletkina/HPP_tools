# -*- coding: utf-8 -*-
__title__ = "Wandart"
__author__ = "olga.poletkina@hpp.com"
__doc__ = """
Author: olga.poletkina@hpp.com
Date: 18.07.2023
___________________________________________________________
Description:
The script applies the host wall material to the door parameter.
___________________________________________________________
How-to:
Press the button.
___________________________________________________________
Prerequisite:
The wall materials must be named in accordance with the 
HPP naming convention!
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

def unit_conventer(
        doc,
        value,
        to_internal=False,
        unit_type=DB.SpecTypeId.Length,
        number_of_digits=None):
    display_units = doc.GetUnits().GetFormatOptions(unit_type).GetUnitTypeId()
    method = DB.UnitUtils.ConvertToInternalUnits if to_internal \
        else DB.UnitUtils.ConvertFromInternalUnits
    if number_of_digits is None:
        return method(value, display_units)
    elif number_of_digits > 0:
        return round(method(value, display_units), number_of_digits)
    return int(round(method(value, display_units), number_of_digits))

doors = [door for door in FEC(doc).OfCategory(
    DB.BuiltInCategory.OST_Doors).WhereElementIsNotElementType().ToElements() if door.Parameter[
    DB.BuiltInParameter.PHASE_DEMOLISHED].AsDouble() == 0]

with DB.Transaction(doc, 'Host material parameter application') as t:
    t.Start()
    material = []
    for door in doors:
        host_mat_parameter = door.LookupParameter('H_TÜ_Wandart')
        if host_mat_parameter == None:
            print('Please apply "H_TÜ_Wandart" parameter!')
            break
        elif host_mat_parameter != None:
            wall = door.Host
            if wall.CurtainGrid == None:
                material_id = wall.GetMaterialIds(False)
                for item in material_id:
                    material = doc.GetElement(item).Name
                    str_to_apply = material.split('_')
            try:
                host_mat_parameter.Set(str_to_apply[2])
            except:
                pass
    t.Commit()

if door.LookupParameter('H_TÜ_Wandart') != None:
    print("The door 'H_TÜ_Wandart' parameter is filled!")