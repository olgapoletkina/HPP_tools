# -*- coding: utf-8 -*-
__title__ = "Location"
__author__ = "olga.poletkina@hpp.com"
__doc__ = """
Author: olga.poletkina@hpp.com
Date: 13.10.2023
___________________________________________________________
Description:
The script checks the location of specific elements in a 
Revit project and creates a filter called "Location check" 
in the active view. The filter includes elements that do 
not match their assigned level parameter based on their 
bounding box elevation. The script applies graphic overrides 
to these elements, highlighting them with a specified 
background pattern and color.
___________________________________________________________
How-to:
Press the button.
___________________________________________________________
Prerequisite:
A proper level types and heights should be applied. 
The script is applicable for projects with level heights 
approximately equal to or greater than 3 meters.
Please limit the visibility or selection of elements to a 
specific area or subset before running the script. 
Running the script on the entire file may 
cause performance issues!
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

uiapp = __revit__
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

# doc = DocumentManager.Instance.CurrentDBDocument
# uiapp = DocumentManager.Instance.CurrentUIApplication
# app = uiapp.Application
# uidoc = uiapp.ActiveUIDocument

def flatten(element, flat_list=None):
    if flat_list is None:
        flat_list = []
    if hasattr(element, "__iter__"):
        for item in element:
            flatten(item, flat_list)
    else:
        flat_list.append(element)
    return flat_list

def group_by_key(elements, key_type='Type'):
    elements = flatten(elements)
    element_groups = {}
    for element in elements:
        if key_type == 'Type':
            key = type(element)
        elif key_type == 'Category':
            for key in DB.BuiltInCategory.GetValues(DB.BuiltInCategory):
                if int(key) == element.Category.Id.IntegerValue:
                    break
        else:
            key = 'Unknown Key'
        if key not in element_groups:
            element_groups[key] = []
        element_groups[key].append(element)
    return element_groups

def get_3d_view_type_id(doc):
    collector = FEC(doc).OfClass(DB.ViewFamilyType)
    for view_type in collector:
        if view_type.ViewFamily == DB.ViewFamily.ThreeDimensional:
            return view_type.Id
    return None

def merge_bounding_boxes(bboxes):
    # type: (list) -> DB.BoundingBoxXYZ
    """Merges multiple bounding boxes"""
    merged_bb = DB.BoundingBoxXYZ()
    merged_bb.Min = DB.XYZ(
        min(bboxes, key=lambda bb: bb.Min.X).Min.X,
        min(bboxes, key=lambda bb: bb.Min.Y).Min.Y,
        min(bboxes, key=lambda bb: bb.Min.Z).Min.Z
    )
    merged_bb.Max = DB.XYZ(
        max(bboxes, key=lambda bb: bb.Max.X).Max.X,
        max(bboxes, key=lambda bb: bb.Max.Y).Max.Y,
        max(bboxes, key=lambda bb: bb.Max.Z).Max.Z
    )
    return merged_bb

def view_exists(doc, view_name):
    views = FEC(doc).OfClass(DB.View).ToElements()
    for view in views:
        if view.Name == view_name:
            return True
    return False

def get_parameter_value_v2(parameter):
    if isinstance(parameter, DB.Parameter):
        storage_type = parameter.StorageType
        if storage_type:
            exec 'parameter_value = parameter.As{}()'.format(storage_type)
            return parameter_value

def unit_conventer(
        doc,
        value,
        to_internal=False,
        unit_type=DB.SpecTypeId.Length,
        number_of_digits=None):
    display_units = doc.GetUnits().GetFormatOptions(unit_type).GetUnitTypeId()
    method = DB.UnitUtils.ConvertToInternalUnits if to_internal \
        else DB.UnitUtils.ConvertFromInternalUnits
    if number_of_digits is None:
        return method(value, display_units)
    elif number_of_digits > 0:
        return round(method(value, display_units), number_of_digits)
    return int(round(method(value, display_units), number_of_digits))

def check_intersection(bbox, element):
    """Check if an element's bounding box intersects with the given bounding box."""
    outline = DB.Outline(bbox.Min, bbox.Max)
    filter_intersects = DB.BoundingBoxIntersectsFilter(outline, False)
    if element.get_BoundingBox(None) is not None:
        return filter_intersects.PassesFilter(doc, element.Id)
    return False

# filter list
categories_filter = List[DB.BuiltInCategory]()
categories_filter.Add(DB.BuiltInCategory.OST_Doors)
categories_filter.Add(DB.BuiltInCategory.OST_Walls)
categories_filter.Add(DB.BuiltInCategory.OST_Ceilings)
categories_filter.Add(DB.BuiltInCategory.OST_Windows)
categories_filter.Add(DB.BuiltInCategory.OST_Furniture)
categories_filter.Add(DB.BuiltInCategory.OST_FurnitureSystems)
categories_filter.Add(DB.BuiltInCategory.OST_Stairs)
categories_filter.Add(DB.BuiltInCategory.OST_StructuralColumns)
# categories_filter.Add(DB.BuiltInCategory.OST_StructuralFoundation)
categories_filter.Add(DB.BuiltInCategory.OST_StructuralFraming)
# categories_filter.Add(DB.BuiltInCategory.OST_Floors)
categories_filter.Add(DB.BuiltInCategory.OST_Railings)
categories_filter.Add(DB.BuiltInCategory.OST_CurtainWallPanels)
categories_filter.Add(DB.BuiltInCategory.OST_CurtainWallMullions)

# create multifilter
multi_category_filter = DB.ElementMulticategoryFilter(categories_filter)

# elements from Active View
elements_list = FEC(doc, doc.ActiveView.Id).WherePasses(multi_category_filter) \
    .WhereElementIsNotElementType().ToElements()

# level parameters
level_parameter_list = [
    DB.BuiltInParameter.FAMILY_LEVEL_PARAM,
    DB.BuiltInParameter.WALL_BASE_CONSTRAINT,
    DB.BuiltInParameter.LEVEL_PARAM,
    DB.BuiltInParameter.SCHEDULE_LEVEL_PARAM,
    DB.BuiltInParameter.STAIRS_BASE_LEVEL_PARAM,
    DB.BuiltInParameter.FAMILY_BASE_LEVEL_PARAM,
    DB.BuiltInParameter.STAIRS_RAILING_BASE_LEVEL_PARAM
]

# create 3d view for bboxes application
view_family_type_id = get_3d_view_type_id(doc)

if view_family_type_id:
    if not view_exists(doc, "Bounding Box View"):
        with DB.Transaction(doc, "Create 3D View") as t:
            t.Start()
            view_3d = DB.View3D.CreateIsometric(doc, view_family_type_id)
            view_3d.Name = "Bounding Box View"
            view_3d.IsSectionBoxActive = False
            t.Commit()
    else:
        pass

# get the view
view_3d = [view for view in FEC(doc).OfClass(DB.View).ToElements() if view.Name == "Bounding Box View"].pop()

# pick all levels from project, that are marked as 'Building Story'
levels = [level for level in FEC(doc).OfClass(DB.Level).WhereElementIsNotElementType().ToElements() \
          if level.Parameter[DB.BuiltInParameter.LEVEL_IS_BUILDING_STORY].AsInteger() == 1]

# sort 'Building Story' floors by elevation
levels_sorted = sorted(levels, key=lambda level: level.Elevation)

# get Bounding boxes by each level
bounding_boxes_by_level = []

# collect elements that are not intersecting the bbox of referenced level
location_check = []
with ProgressBar(cancellable=True) as pb:
    for [index, level], counter in zip(enumerate(levels_sorted), range(len(levels_sorted))):
        if index < len(levels_sorted) - 1:
            bbox = merge_bounding_boxes([level.get_BoundingBox(view_3d), \
                                        levels_sorted[index+1].get_BoundingBox(view_3d)])
        else:
            distance = unit_conventer(doc, 5)
            bbox = level.get_BoundingBox(view_3d)
            bbox.Max = DB.XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z + distance)
        for element in elements_list:
            for level_parameter in level_parameter_list:
                if element.Parameter[level_parameter] is not None:
                    level_param = element.Parameter[level_parameter]
                    if get_parameter_value_v2(level_param) == level.Id:

                        # now check if bbox intersects the element
                        if check_intersection(bbox, element) == False:
                            location_check.append(element)
        if pb.cancelled:
            break
        else:
            pb.update_progress(counter, len(levels_sorted))
            bounding_boxes_by_level.append(bbox)

i_collection = List[DB.ElementId]()

for element in location_check:
    i_collection.Add(element.Id)

# create filter
active_view = doc.ActiveView
graphic_settings = DB.OverrideGraphicSettings()
graphic_settings.SetSurfaceBackgroundPatternVisible(True)
patterns = FEC(doc).OfClass(DB.FillPatternElement).ToElements()
solid_pattern = [pattern for pattern in patterns if pattern.GetFillPattern().IsSolidFill]
graphic_settings.SetSurfaceBackgroundPatternId(solid_pattern[0].Id)
graphic_settings.SetSurfaceBackgroundPatternColor(DB.Color(255, 0, 0))
graphic_settings.SetSurfaceForegroundPatternColor(DB.Color(255, 0, 0))

with DB.Transaction(doc, 'New filter') as t:
    t.Start()
    for filter in FEC(doc).OfClass(DB.SelectionFilterElement).ToElements():
        if doc.GetElement(filter.Id).Name == 'Location check':
            doc.Delete(filter.Id)
        else:
            pass
    filter = DB.SelectionFilterElement.Create(
        doc,
        'Location check'
    )
    filter.AddSet(
        i_collection
    )
    active_view.AddFilter(filter.Id)
    active_view.SetFilterOverrides(filter.Id, graphic_settings)
    t.Commit()

print('Filter "Location check" is created')
