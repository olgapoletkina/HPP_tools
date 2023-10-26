# -*- coding: utf-8 -*-
__title__ = "Title Blocks"
__author__ = "olga.poletkina@hpp.com"
__doc__ = """
Author: olga.poletkina@hpp.com
Date: 01.08.2023
___________________________________________________________
Description:
The script retrieves title block families and removes 
them from the project if the document is not workshared.
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

project_families = FEC(doc).OfClass(DB.Family).ToElements()
title_blocks = []
title_blocks_names = []
for item in project_families:
    category = DB.Category.GetCategory(doc, item.FamilyCategoryId)
    if item.FamilyCategory.Id == DB.ElementId(DB.BuiltInCategory.OST_TitleBlocks):
        title_blocks.append(item.Id)
        title_blocks_names.append(item.Name)

if doc.IsWorkshared == True:
    print('you need to detach file first')
elif len( title_blocks) > 0:
    print('The following Title block families were removed:')
    for name in title_blocks_names:
        print(name)
    with DB.Transaction(doc, 'Delete Title Blocks') as t:
        t.Start()
        for item_id in title_blocks:
            try:
                doc.Delete(item_id)
            except:
                pass
        t.Commit()
else:
    print('No Title Block families were detected!')
    # print(FEC(doc).OfCategory(DB.BuiltInCategory.OST_TitleBlocks).ToElements())