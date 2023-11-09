# -*- coding: utf-8 -*-
__title__ = "DWG"
__author__ = "olga.poletkina@hpp.com"
__doc__ = """
Author: olga.poletkina@hpp.com
Date: 01.08.2023
___________________________________________________________
Description:
This script identifies CAD link elements (DWG files) in 
the project and removes them if the document is not 
workshared. It then outputs a message listing the names 
of the deleted DWG files, if any were found.
___________________________________________________________
How-to:
Press the button.
___________________________________________________________
Prerequisite:
Detatch the file from the central.
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

links = FEC(doc).OfClass(DB.CADLinkType).ToElements()

if doc.IsWorkshared == True:
    print('you need to detach file first')
elif len(links) > 0:
    print('The following DWG files were removed:')
    for link in links:
        print(link.Parameter[DB.BuiltInParameter.ALL_MODEL_TYPE_NAME].AsString())
    with DB.Transaction(doc, 'Delete all DWG') as t:
        t.Start()
        for item in links:
            try:
                doc.Delete(item.Id)
            except:
                pass
        t.Commit()
else:
    print('No DWG were detected!')