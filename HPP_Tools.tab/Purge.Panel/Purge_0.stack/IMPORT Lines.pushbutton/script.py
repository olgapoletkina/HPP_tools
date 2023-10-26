# -*- coding: utf-8 -*-
__title__ = "IMPORT Lines"
__author__ = "olga.poletkina@hpp.com"
__doc__ = """
Author: olga.poletkina@hpp.com
Date: 18.07.2023
___________________________________________________________
Description:
This script removes imported line types from the Revit 
project. It collects line patterns with "IMPORT" in 
their names, deletes them from the project, 
and provides a list of removed line types.
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

clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
clr.AddReference('RevitNodes')
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector as FEC
from Autodesk.Revit.UI import Selection

from System import *

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

lines_to_remove = []
lines_to_remove_names = []
line_patterns = FEC(doc).OfClass(DB.LinePatternElement).ToElements()
for line_pattern in line_patterns:
    if 'IMPORT' in line_pattern.Name:
        lines_to_remove.append(line_pattern)
        lines_to_remove_names.append(line_pattern.Name)

if len(lines_to_remove) > 0:
    with DB.Transaction(doc, 'Delete Import lines') as t:
        t.Start()
        for item in lines_to_remove:
            try:
                doc.Delete(item.Id)
            except:
                pass
        t.Commit()
    print('The following line types were removed:')
    for name in lines_to_remove_names:
        print(name)
else:
    print('No IMPORTED lines were detected!')