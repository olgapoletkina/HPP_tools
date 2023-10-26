# -*- coding: utf-8 -*-
__title__ = "Glass Area"
__author__ = "olga.poletkina@hpp.com"
__doc__ = """
Author: olga.poletkina@hpp.com
Date: 18.07.2023
___________________________________________________________
Description:
This script calculates the area of the outside glassing 
by extracting glass elements from outside windows 
and curtain walls that are marked as "HA" 
or match a wall type parameter "Exterior".
___________________________________________________________
How-to:
Press the button.
___________________________________________________________
Prerequisite:
The wall materials must be named in accordance with the 
HPP naming convention!
The corresponding parameter should be applied!
Please note that window families should not have 
double glassing, glassing elements should not intersect 
with the frame, and all elements should be marked with 
a proper subcategory. Also, be sure to use proper 
Interior/Exterior parameter/naming for the elements.
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

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def get_all_solids(element, g_options, solids=None):
    '''retrieve all solids from elements'''
    if solids is None:
        solids = []
    if hasattr(element, "Geometry"):
        for item in element.Geometry[g_options]:
            get_all_solids(item, g_options, solids)
    elif isinstance(element, DB.GeometryInstance):
        for item in element.GetInstanceGeometry():
            get_all_solids(item, g_options, solids)
    elif isinstance(element, DB.Solid):
        solids.append(element)
    elif isinstance(element, DB.FamilyInstance):
        for item in element.GetSubComponentIds():
            family_instance = element.Document.GetElement(item)
            get_all_solids(family_instance, g_options, solids)
    return solids

def flatten(element, flat_list=None):
    if flat_list is None:
        flat_list = []
    if hasattr(element, "__iter__"):
        for item in element:
            flatten(item, flat_list)
    else:
        flat_list.append(element)
    return flat_list

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

g_options = DB.Options()

# collect all CW panels
curtain_panels = FEC(doc).OfCategory(DB.BuiltInCategory.OST_CurtainWallPanels).WhereElementIsNotElementType().ToElements()

# calculating area by extracting area parameter from panel
panels_area = 0
for panel in curtain_panels:
    if hasattr(panel, 'Symbol'):
        if panel.Symbol.Parameter[DB.BuiltInParameter.MATERIAL_ID_PARAM] != None:
            panel_material_id = panel.Symbol.Parameter[DB.BuiltInParameter.MATERIAL_ID_PARAM].AsElementId()
            material = doc.GetElement(panel_material_id)
            if material != None:
                material_name = material.Name
                host_wall = panel.Host
                if host_wall.WallType.Parameter[DB.BuiltInParameter.FUNCTION_PARAM].AsValueString() == 'Exterior':
                    if 'GLA' in material_name:
                        glass_area = unit_conventer(doc, 
                                                    panel.Parameter[DB.BuiltInParameter.HOST_AREA_COMPUTED].AsDouble(),
                                                    unit_type=DB.SpecTypeId.Area) 
                        panels_area += glass_area

# collect all windows
windows = FEC(doc).OfCategory(DB.BuiltInCategory.OST_Windows).WhereElementIsNotElementType().ToElements()

# calculating area using the largest face of a solid
# filtering only outside windows
solids = []
for window in windows:
    if 'HA' in window.Symbol.Family.Name:
        w_solids = get_all_solids(window, g_options)
        for w_solid in w_solids:
            mat_id = w_solid.GraphicsStyleId
            mat = doc.GetElement(mat_id)
            if mat != None:
                if 'las' in mat.GraphicsStyleCategory.Name:
                    solids.append(w_solid)

# calculating the area
solids_areas = []
for solid in solids:
    faces = solid.Faces
    area_list = []
    for face in faces:
        area_list.append(face.Area)
    solids_areas.append(sorted(area_list)[-1])

# converting to m2
windows_area = (unit_conventer(
    doc, 
    sum(solids_areas),
    unit_type=DB.SpecTypeId.Area
))

print('Total facade glass area is: {}'.format(panels_area + windows_area))
print('***')
print('The glass area of following types were used in calculation:')

# printing the families we used for area calculation
final_list = []
for window in windows:
    if 'HA' in window.Symbol.Family.Name:
        if window.Symbol.Family.Name not in final_list:
            final_list.append(window.Symbol.Family.Name)
for panel in curtain_panels:
    if hasattr(panel, 'Host'):
        if panel.Host.WallType.Parameter[DB.BuiltInParameter.FUNCTION_PARAM].AsValueString() == 'Exterior':
            if 'las' in panel.Name:
                if panel.Host.Name not in final_list:
                    final_list.append(panel.Host.Name)

for item in final_list:
    print(item)