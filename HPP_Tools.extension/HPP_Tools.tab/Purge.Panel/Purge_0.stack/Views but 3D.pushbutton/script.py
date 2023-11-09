# -*- coding: utf-8 -*-
__title__ = "Views but 3D"
__author__ = "olga.poletkina@hpp.com"
__doc__ = """
Author: olga.poletkina@hpp.com
Date: 01.08.2023
___________________________________________________________
Description:
The script allows the user to select specific 3D views 
from the current Revit project. It then deletes all other 
views, excluding the selected 3D views, provided that the 
project is not workshared and at least 
one 3D view is selected.
___________________________________________________________
How-to:
- Pick the 3D view
- Press the button.
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
from System import *

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

view_selected = uidoc.Selection.GetElementIds()

views = FEC(doc).OfClass(DB.View).ToElements()
delete = []
keep = []
for view in views:
    if view.ViewType != ViewType.ThreeD:
        delete.append(view.Id)
    elif view.Id not in view_selected:
        delete.append(view.Id)
    else:
        keep.append(view)

if doc.IsWorkshared == True:
    print('you need to detach file first')
elif view_selected == None:
    print('pick 3D view(s)')
elif len(keep) >= 1:
    with DB.Transaction(doc, 'Delete Views') as t:
        t.Start()
        for view_id in delete:
            try:
                doc.Delete(view_id)
            except:
                pass
        t.Commit()
    print('views are removed')
else:
    Exception
    print('pick 3D view(s)')