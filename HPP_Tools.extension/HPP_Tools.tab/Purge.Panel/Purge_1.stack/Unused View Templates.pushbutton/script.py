# -*- coding: utf-8 -*-
__title__ = "Unused View Templates"
__author__ = "olga.poletkina@hpp.com"
__doc__ = """
Author: olga.poletkina@hpp.com
Date: 01.08.2023
___________________________________________________________
Description:
The script identifies unused View Templates in the project 
and removes them if the document is not workshared. 
It then outputs a message listing the names of the 
deleted templates, if any were found.
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
from System.Collections.Generic import List

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

views = FEC(doc).OfClass(DB.View)
applied_templates_id = [view.ViewTemplateId for view in views]
templates_id = [view.Id for view in views if view.IsTemplate == True]

templates_to_delete = []
templates_to_delete_names = []
for template_id in templates_id:
    if template_id not in applied_templates_id:
        templates_to_delete.append(template_id)
        templates_to_delete_names.append(doc.GetElement(template_id).Name)

if doc.IsWorkshared == True:
    print('you need to detach file first')
elif len(templates_to_delete) > 0:
    with DB.Transaction(doc, 'Delete Title Blocks') as t:
        t.Start()
        for item_id in templates_to_delete:
            try:
                doc.Delete(item_id)
            except:
                pass
        t.Commit()
    print('The following templates were deleted:')
    for name in templates_to_delete_names:
        print(name)
else:
    print('No unused View Templates were detected!')
