# -*- coding: utf-8 -*-
__title__ = "Room Height"
__author__ = "olga.poletkina@hpp.com"
__doc__ = """
Author: olga.poletkina@hpp.com
Date: 18.07.2023
___________________________________________________________
Description:
This script performs several tasks related to room heights
in a Revit project. It excludes staircases from the room 
list, calculates volumes for rooms, adjusts room heights, 
calculates heights based on floors and ceilings, applies 
height parameters to rooms, creates a schedule to check 
room heights, and provides instructions for reviewing 
the heights in the generated schedule 
named "Raumhöhe Check.".
___________________________________________________________
How-to:
Press the button.
___________________________________________________________
Prerequisite:
Not applicable to the rooms with a complex ceiling geometry.
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
from Autodesk.Revit.DB import *
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector as FEC

from System import *

import pyrevit
from pyrevit.forms import ProgressBar

from Snippets._functions import unit_conventer, get_room_boundary

# uiapp = __revit__
doc = __revit__.ActiveUIDocument.Document
# uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

# excluding staircases from the room list
rooms = []
for room in FEC(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).ToElements():
    if room.Parameter[DB.BuiltInParameter.ROOM_PHASE].AsValueString() == 'Neu':
        if 'Treppenhaus' not in room.Parameter[DB.BuiltInParameter.ROOM_NAME].AsString() and \
            'Treppe' not in room.Parameter[DB.BuiltInParameter.ROOM_NAME].AsString() and \
            'TH' not in room.Parameter[DB.BuiltInParameter.ROOM_NAME].AsString() and \
            'TRH' not in room.Parameter[DB.BuiltInParameter.ROOM_NAME].AsString() and \
            'chacht' not in room.Parameter[DB.BuiltInParameter.ROOM_NAME].AsString():
            rooms.append(room)

# for room in rooms:
#     print(room.Parameter[DB.BuiltInParameter.ROOM_NAME].AsString())

# switch on the volume calculation in rooms
with DB.Transaction(doc, 'Set area and volume calculation to True') as t:
    t.Start()
    settings = AreaVolumeSettings.GetAreaVolumeSettings(doc)
    settings.ComputeVolumes = True
    settings.SetSpatialElementBoundaryLocation(SpatialElementBoundaryLocation.Finish, SpatialElementType.Room)  
    t.Commit()

# apply the height to room, so it will cross the upper border
with DB.Transaction(doc, 'Change room height') as t:
    t.Start()
    for room in rooms:
        height_parameter = room.Parameter[DB.BuiltInParameter.ROOM_UPPER_OFFSET]
        height_parameter.Set(unit_conventer(doc, 7, to_internal=True))
    t.Commit()

options = SpatialElementBoundaryOptions()
boundloc = AreaVolumeSettings.GetAreaVolumeSettings(doc).GetSpatialElementBoundaryLocation(SpatialElementType.Room)
options.SpatialElementBoundaryLocation = boundloc

# get all bounding room geometries
element_list = []
curve_list = []
room_element_list = []
with ProgressBar(cancellable=True) as pb:
    try:
        error_report = None
        if isinstance(rooms, list):
            for room, counter in zip(rooms, range(len(rooms))):
                element_list.append(get_room_boundary(doc, room, options)[0])
                curve_list.append(get_room_boundary(doc, room, options)[1])
                room_element_list.append([room, get_room_boundary(doc, room, options)[0]])
                pb.update_progress(counter, len(rooms))
                
                if pb.cancelled:
                    cancelled = True
                    break
        else:
            element_list = get_room_boundary(doc, rooms, options)[0]
            curve_list = get_room_boundary(doc, rooms, options)[1]
            room_element_list.append([room, element_list])
    except:
    # if error accurs anywhere in the process catch it
        import traceback
        error_report = traceback.format_exc()

# sort only floors and ceilings
floors_ceilings = []

for room in rooms:
    if room.LookupParameter('H_RA_lichte_Höhe') == None:
        print('Please apply parameter "H_RA_lichte_Höhe" to room!')
        break

# for room in rooms:
#     if room.LookupParameter('H_RA_lichte_Höhe') == None:
#         raise Exception ('Please apply parameter "H_RA_lichte_Höhe" to rooms!')

cancelled = False

with DB.Transaction(doc, 'Apply heights') as t:
    t.Start()

    with ProgressBar(cancellable=True) as pb:
    # for element, counter in zip(elements, range(0, len(elements))):
        for i in range(len(room_element_list)):
            # get room parameters
            if room_element_list[i][0].LookupParameter('H_RA_lichte_Höhe') != None:
                l_H_parameter = room_element_list[i][0].LookupParameter('H_RA_lichte_Höhe')
                room_height_parameter = room_element_list[i][0].Parameter[DB.BuiltInParameter.ROOM_UPPER_OFFSET]
                
                # variable for collecting ceilings and floors only
                f_c = []
                cat_name = []
                for elem in room_element_list[i][1]:
                    if elem != None:
                        cat_name.append(elem.Category.Name)
                if 'Roofs' not in cat_name:
                    for el in room_element_list[i][1]:
                        if el != None:
                            if el.Category.Name == 'Floors' or el.Category.Name == 'Ceilings':
                                f_c.append(el)
                # take only rooms that have more then one ceiling/floor element
                    if len(f_c)>1:
                # list contains room Id and upper/lower border
                        floors_ceilings.append([room_element_list[i][0].Id, f_c])
                        # print(floors_ceilings)
                        height_sum = 0
                        count = 0
                        upper_border = []
                        lower_border = []
                        room_h_to_apply = []
                        room_h_number = []
                # calculate the heights
                        for element in f_c:
                            count += 1
                # rules to calculate height for floor type. if floor belongs to room level
                            if element.Category.Name == 'Floors':
                                if room_element_list[i][0].Parameter[
                                    DB.BuiltInParameter.ROOM_LEVEL_ID].AsElementId() == element.Parameter[
                                    DB.BuiltInParameter.LEVEL_PARAM].AsElementId():
                                    par_height = element.Parameter[DB.BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM].AsDouble()
                # and if floor doesn`t belong to room level
                                else:
                                    par_height = element.Parameter[
                                        DB.BuiltInParameter.STRUCTURAL_ELEVATION_AT_TOP].AsDouble() - doc.GetElement(
                                        room_element_list[i][0].Parameter[
                                        DB.BuiltInParameter.ROOM_LEVEL_ID].AsElementId()).Parameter[
                                        DB.BuiltInParameter.LEVEL_ELEV].AsDouble()
                # height from ceiling
                            elif element.Category.Name == 'Ceilings':
                                if element.Parameter[DB.BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM] != None:
                                    par_height = element.Parameter[DB.BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM].AsDouble()
                # separate floors and ceilings for upper and lower border by camparing with a mean height
                            height_sum += par_height
                            middle = height_sum / count
                            # print(middle)
                            if par_height > middle:
                                if element.Category.Name == 'Floors':
                                    upper_border.append(par_height - element.Parameter[
                                        DB.BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM].AsDouble())
                                    # print(upper_border)
                                else:
                                    upper_border.append(par_height)
                            else:
                                lower_border.append(par_height)
                        # print(room_element_list[i][0].Id)
                        # print(upper_border, lower_border)
                # variable max_height to apply to room limit border
                        max_height = None
                # applying parameters only for rooms with one lower border element
                        if len(lower_border) == 1:
                            for up_element in upper_border:
                                room_height = unit_conventer(doc,(up_element - lower_border[0]))
                                if str(room_height) not in room_h_to_apply:
                                    room_h_to_apply.append(str(room_height))
                                room_h_number.append(up_element - lower_border[0])
                                max_height = max(room_h_number)
                # re-apply room height param
                                """
                                Script needs changes in case several level combination on one floor to be used!
                                Parameter Limit Offset, room_height_parameter.Set(max_height) should be set to Zero!
                                """
                                room_height_parameter.Set(max_height)
                # apply room heights in text param
                        if l_H_parameter != None:
                            string_param = ', '.join(room_h_to_apply)
                            l_H_parameter.Set(string_param)
            pb.update_progress(i, len(room_element_list))
        
            if pb.cancelled:
                cancelled = True
                break

    if cancelled:
        print('Operation is cancelled!')
    t.Commit()

rooms_no_parameter_applied = []

for room in rooms:
    if room.LookupParameter('H_RA_lichte_Höhe') != None:
        if room.LookupParameter('H_RA_lichte_Höhe').AsString() == '' or room.LookupParameter('H_RA_lichte_Höhe').AsString() == None:
            rooms_no_parameter_applied.append(room)

# create schedule
if len(rooms_no_parameter_applied) > 0:
    with DB.Transaction(doc, 'Create list "Raumhöhe Check"') as t:
        t.Start()
        if len([item for item in FEC(doc).OfClass(DB.ViewSchedule).ToElements() if item.Name == 'Raumhöhe Check']) > 0:
            print('Existing schedule "Raumhöhe Check" is modified')
        else:
            room_schedule = DB.ViewSchedule.CreateSchedule(
                            doc,
                            DB.ElementId(DB.BuiltInCategory.OST_Rooms)
                        )
            room_schedule.Name = 'Raumhöhe Check'
            s_definition = room_schedule.Definition
            # add schedule fields
            room_field = s_definition.AddField(
               DB.ScheduleFieldType.Instance,
                DB.ElementId(
                    DB.BuiltInParameter.ROOM_NAME
                )
            )
            room_param_id = []
            for room in rooms:
                if room.LookupParameter('H_RA_lichte_Höhe') != None:
                    room_param_id.append(room.LookupParameter('H_RA_lichte_Höhe').Id)
                else:
                    pass
            fussboden_field = s_definition.AddField(
               DB.ScheduleFieldType.Instance,
                room_param_id[0]
            )

        t.Commit()
else:
    if room.LookupParameter('H_RA_lichte_Höhe') != None:
        print('Parameter Raumhöhe is applied.')
if room.LookupParameter('H_RA_lichte_Höhe') != None:
    print('Please check room heights in schedule "Raumhöhe Check"!')

"""
Script needs changes in case several level combination on one floor to be used!
Parameter Limit Offset, room_height_parameter.Set(max_height) should be set to Zero!
"""


