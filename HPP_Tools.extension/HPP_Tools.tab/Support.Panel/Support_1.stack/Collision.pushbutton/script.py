# -*- coding: utf-8 -*-
__title__ = "Collision"
__author__ = "olga.poletkina@hpp.com"
__doc__ = """
Author: olga.poletkina@hpp.com
Date: 13.10.2023
___________________________________________________________
Description:
This tool checks whether the Kollisionkörper of a door or 
window family intersects with other geometries in the 
project. Based on that, a schedule containing the 
intersected elements is created or modified.
___________________________________________________________
How-to:
Press the button.
___________________________________________________________
Prerequisite:
The corresponding parameters 'H_TÜ_Kollisionskörper 
einschalten' and 'H_OQ_Kollisionsprüfung' should be applied!
Please limit the visibility or selection of elements to a 
specific area or subset before running the script. 
Running script on the entire file may cause performance 
issues!
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

import pyrevit
from pyrevit.forms import ProgressBar

from Snippets._functions import flatten, to_list

# uiapp = __revit__
doc = __revit__.ActiveUIDocument.Document
# uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

# categories filter list
categories_to_check = List[DB.BuiltInCategory]()
categories_to_check.Add(DB.BuiltInCategory.OST_Walls)
categories_to_check.Add(DB.BuiltInCategory.OST_Floors)
categories_to_check.Add(DB.BuiltInCategory.OST_Ceilings)
categories_to_check.Add(DB.BuiltInCategory.OST_Furniture)
categories_to_check.Add(DB.BuiltInCategory.OST_FurnitureSystems)
categories_to_check.Add(DB.BuiltInCategory.OST_Lights)
categories_to_check.Add(DB.BuiltInCategory.OST_Stairs)

# create multifilter
multi_category_filter = DB.ElementMulticategoryFilter(categories_to_check)

# door and window filter
door_window_category = List[DB.BuiltInCategory]()
door_window_category.Add(DB.BuiltInCategory.OST_Doors)
door_window_category.Add(DB.BuiltInCategory.OST_Windows)
door_window_filter = DB.ElementMulticategoryFilter(door_window_category)

# collecting doors and windows elements
elements = FEC(doc).WherePasses(door_window_filter).WhereElementIsNotElementType().ToElements()

# check if 'H_TÜ_Kollisionskörper einschalten' is in the project
shared_parameter_name = 'H_TÜ_Kollisionskörper einschalten'

param_element_id = None
for element in elements:
    for param in element.Parameters:
        if param.Definition is not None and param.Definition.Name == shared_parameter_name:
            param_element_id = param.Id
            break

cancelled = False

clashed_elements = []

if param_element_id == None:
    print("Please apply parameter 'H_TÜ_Kollisionskörper einschalten' to the families.")

else:

    # rule, that checks if family has parameter 'H_TÜ_Kollisionskörper einschalten'
    window_door_contains_rule = DB.ParameterFilterRuleFactory.CreateHasValueParameterRule(
        param_element_id
    )

    # filter collects only windows and doors containing parameter above
    and_rule_filter = DB.LogicalAndFilter(
        to_list(
            [
                door_window_filter,
                DB.ElementParameterFilter(window_door_contains_rule)
            ],
            DB.ElementFilter
        ))

    # filter for categories, dors and windows with 'H_TÜ_Kollisionskörper einschalten'
    or_rule_filter = DB.LogicalOrFilter(
        to_list(
            [
                multi_category_filter,
                and_rule_filter
            ],
            DB.ElementFilter
        ))

    # collect elements from active view
    elements_check_list = FEC(doc, doc.ActiveView.Id).WherePasses(or_rule_filter).WhereElementIsNotElementType().ToElements()

    # get filters
    filters_list = [DB.ElementIntersectsElementFilter(el) for el in elements_check_list]

    # get collisions and initiate the progress bar
    total_result = []
    with ProgressBar(cancellable=True) as pb:
        for filter, counter in zip(filters_list, range(0, len(filters_list))):
            result = [
                doc.GetElement(el_id) for el_id in FEC(doc).WherePasses(or_rule_filter).WherePasses(filter).ToElementIds()
                ]
            if pb.cancelled:
                cancelled = True
                break
            else:
                pb.update_progress(counter, len(filters_list))
            total_result.append(result)

    # unique clash elements by id
    clashed_elements_ids = []
    for element in flatten(total_result):
        if element.Category.Id == DB.Category.GetCategory(
            doc, DB.BuiltInCategory.OST_Doors).Id or element.Category.Id == DB.Category.GetCategory(
            doc, DB.BuiltInCategory.OST_Windows).Id:
            if element.Id not in clashed_elements_ids:
                clashed_elements_ids.append(element.Id)

    # get clashed elements
    clashed_elements = flatten([doc.GetElement(el_id) for el_id in clashed_elements_ids])

    # get not clashed elements ids
    not_clashed_elements_ids = [el.Id for el in FEC(doc).WherePasses(
        door_window_filter).WhereElementIsNotElementType().ToElements(
        ) if el.Id not in clashed_elements_ids]

    # get clashed elements
    not_clashed_elements = flatten([doc.GetElement(el_id) for el_id in not_clashed_elements_ids])
    for item in clashed_elements:
        if item.LookupParameter('H_OQ_Kollisionsprüfung') == None:
            print('Please add "H_OQ_Kollisionsprüfung" parameter to the project!')
            break

    # apply parameters
    try:
        with DB.Transaction(doc, 'Assign Collision Check Parameter') as t:
            t.Start()
            for item in clashed_elements:
                parameter = item.LookupParameter('H_OQ_Kollisionsprüfung')
                parameter.Set(True)
            for item in not_clashed_elements:
                parameter = item.LookupParameter('H_OQ_Kollisionsprüfung')
                parameter.Set(False)
            t.Commit()
    except:
        Exception


if cancelled:
    print('Operation is cancelled!')
else:
    # schedule creation
    if len(clashed_elements) > 0:
        try:
            with DB.Transaction(doc, 'Create or Modify Schedule') as t:
                t.Start()
                if len([item for item in FEC(doc).OfClass(DB.ViewSchedule).ToElements() if item.Name == 'Collision Check']) > 0:
                    print('Existing schedule "Collision Check" is modified')
                else:
                    multi_schedule = DB.ViewSchedule.CreateSchedule(
                        doc,
                        DB.ElementId(DB.BuiltInCategory.INVALID)
                    )
                    multi_schedule.Name = 'Collision Check'
                s_definition = multi_schedule.Definition

                # add schedule fields
                family_field = s_definition.AddField(
                DB.ScheduleFieldType.Instance,
                    DB.ElementId(
                        DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM
                    )
                )
                collision_param_id = clashed_elements[0].LookupParameter('H_OQ_Kollisionsprüfung').Id
                collision_field = s_definition.AddField(
                DB.ScheduleFieldType.Instance,
                    collision_param_id
                )

                # add schedule filter
                collision_filter = s_definition.AddFilter(
                    DB.ScheduleFilter(
                        collision_field.FieldId, 
                        DB.ScheduleFilterType.HasValue
                    )
                )
                collision_filter_yes = s_definition.AddFilter(
                    DB.ScheduleFilter(
                        collision_field.FieldId, 
                        DB.ScheduleFilterType.Equal,
                        True
                    )
                )
                t.Commit()
            print('Check the schedule named "Collision Check"!')
        except:
            Exception
    elif param_element_id != None:
        print('No clashed doors or windows were detected on active view!')