# -*- coding: utf-8 -*-
__title__ = "Unused Filters"
__author__ = "olga.poletkina@hpp.com"
__doc__ = """
Author: olga.poletkina@hpp.com
Date: 01.08.2023
___________________________________________________________
Description:
The script identifies all filters used in views and then 
removes the unused filters from the project. It outputs a 
message listing the names of the deleted 
filters if any were found.
___________________________________________________________
How-to:
Press the button.
___________________________________________________________
Prerequisite:
None.
___________________________________________________________
"""
import clr
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI")
import Autodesk.Revit
from Autodesk.Revit.Exceptions import InvalidOperationException
from Autodesk.Revit import DB
from Autodesk.Revit.DB import ElementId, FilteredElementCollector as FEC
from System.Collections.Generic import *

clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from RevitServices.Persistence import DocumentManager

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

used_filter_ids_list = []
for view in FEC(doc).OfClass(DB.View).ToElements():
    try:
        filter_ids = view.GetFilters()
        used_filter_ids_list.append(filter_ids)
    except:
        used_filter_ids_list.append(None)

used_filter_ids_list_flatten = []
for item in used_filter_ids_list:
  if item != None and str(item) != 'Empty List':
    for i in item:
      used_filter_ids_list_flatten.append(i)

filter_param = DB.ElementClassFilter(DB.ParameterFilterElement)
filter_select = DB.ElementClassFilter(DB.SelectionFilterElement)
or_filter = DB.LogicalOrFilter(
  filter_param,
  filter_select
)
all_filters = FEC(doc).WherePasses(or_filter).ToElements()

not_used_filters = []
not_used_filter_names = []
for filter in all_filters:
    if filter.Id not in used_filter_ids_list_flatten:
        not_used_filters.append(filter)
        not_used_filter_names.append(filter.Name)

if len(not_used_filters) > 0:
  with DB.Transaction(doc, 'Delete Filters') as t:
        t.Start()
        for item in not_used_filters:
            try:
                doc.Delete(item.Id)
            except:
                pass
        t.Commit()
  print('The following filters were deleted:')
  for name in not_used_filter_names:
    print(name)
else:
  print('No unused filters were detected!')