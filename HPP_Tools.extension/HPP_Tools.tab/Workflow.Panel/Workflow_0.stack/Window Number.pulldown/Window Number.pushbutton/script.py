# -*- coding: utf-8 -*-
__title__ = "Nr. format: F.Width.Sill Height.M"
__author__ = "olga.poletkina@hpp.com"
__doc__ = """
Author: olga.poletkina@hpp.com
Date: 18.07.2023
___________________________________________________________
!!! Please note the format for the number that this 
program will generate: It will start with 'F', followed 
by the width of the window, then the sill height, and end 
with 'M' if the family is mirrored. All similar families 
will receive similar numbers !!!
___________________________________________________________
Description:
This script assigns window numbers based on window 
dimensions and sill height. It checks if the required 
parameters are available and calculates the window number 
accordingly. The generated window numbers 
are printed as output.
___________________________________________________________
How-to:
Press the button.
___________________________________________________________
Prerequisite:
The corresponding parameter should be applied!
___________________________________________________________
"""

import clr
clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
clr.AddReference('RevitNodes')
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)
clr.AddReference('System.Drawing')
from System.Drawing import Color, ColorTranslator

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit import DB, UI
from Autodesk.Revit.DB import FilteredElementCollector as FEC
from Autodesk.Revit.DB import Architecture as AR
from Autodesk.Revit.UI import Selection as SEL
from System import *

from Snippets._functions import unit_converter

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

windows = [window for window in FEC(doc).OfCategory(
    DB.BuiltInCategory.OST_Windows).WhereElementIsNotElementType() if window.Parameter[
    DB.BuiltInParameter.PHASE_CREATED].AsValueString() == 'Neu']

window_numbers = []

with DB.Transaction(doc, 'Assign Window Number') as t:
    t.Start()
    for window in windows:
        if window.Symbol.Parameter[DB.BuiltInParameter.WINDOW_WIDTH] != None and \
            window.Parameter[DB.BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM] != None and \
            window.Parameter[DB.BuiltInParameter.WINDOW_WIDTH] != None:
            
            parameter = window.LookupParameter('H_FE_Fensternummer')
            sill_height = round(unit_converter(
                doc,
                window.Parameter[DB.BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM].AsDouble()
            ) * 100)
            if window.Symbol.get_Parameter(DB.BuiltInParameter.WINDOW_WIDTH).AsDouble() > 0:
                width = round(unit_converter(
                    doc,
                    window.Symbol.Parameter[DB.BuiltInParameter.WINDOW_WIDTH].AsDouble()
                ) * 100)
            else:
                width = round(unit_converter(
                    doc,
                    window.Parameter[DB.BuiltInParameter.WINDOW_WIDTH].AsDouble()
                ) * 100)
            mirrored = window.Mirrored
            if window.Mirrored == True:
                window_number = 'F.' + str(width)[:-2] + '.' + str(sill_height)[:-2] + '.M'
                parameter.Set(window_number)
            else:
                window_number = 'F.' + str(width)[:-2] + '.' + str(sill_height)[:-2]
                parameter.Set(window_number)
            window_numbers.append(window_number)
    t.Commit()

if len(window_numbers) > 0:
    print('Following window numbers were generated:')
    for number in window_numbers:
        print(number)
else:
    print('No numbers were generated')