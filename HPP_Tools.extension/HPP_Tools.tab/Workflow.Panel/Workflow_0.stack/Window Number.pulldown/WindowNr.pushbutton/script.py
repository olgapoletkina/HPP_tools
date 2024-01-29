# -*- coding: utf-8 -*-
__title__ = "Nr. format: RoomNr.F00"
__author__ = "olga.poletkina@hpp.com"
__doc__ = """
Author: olga.poletkina@hpp.com
Date: 29.01.2024
___________________________________________________________
!!! Please note the format for the number that this program 
will generate: It will start with a room number, followed 
by 'F', and then a sequential counting number if several 
doors belong to one room !!!
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

if len(windows) > 0:
    phase_created = [window.Parameter[DB.BuiltInParameter.PHASE_CREATED].AsElementId() for window in windows]
    new_phase = doc.GetElement(phase_created[0])
else:
    print('There are no New Phase window families in the project')

# print(new_phase)

room_window_count = {}

with DB.Transaction(doc, 'Assign Window Number') as t:
    t.Start()
    for window in windows:
        parameter = window.LookupParameter('H_FE_Fensternummer')
        
        if window.ToRoom[new_phase] is not None:
            room_window = window.ToRoom[new_phase].LookupParameter('H_RA_Raumnummer').AsString()
            
            # Check if the room number is already in the dictionary
            if room_window in room_window_count:
                # Increment the count and update window number
                room_window_count[room_window] += 1
                window_number = '{}.F{:02d}'.format(room_window, room_window_count[room_window])
            else:
                # Initialize count and set window number
                room_window_count[room_window] = 0
                window_number = room_window + '.F'

            parameter.Set(window_number)
        else:
            parameter.Set('')
            # options = DB.ExternalDefinitionCreationOptions('H_RA_Raumnummer', DB.ParameterType.Text)
            # options.HideWhenNoValue = True
            # parameter.ClearValue()

            # options.HideWhenNoValue = False
    print('Numbers are generated')
    t.Commit()