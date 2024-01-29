# -*- coding: utf-8 -*-
__title__ = "Nr. format: T.RoomNr.00"
__author__ = "olga.poletkina@hpp.com"
__doc__ = """
Author: olga.poletkina@hpp.com
Date: 29.01.2024
___________________________________________________________
!!! Please note the format for the number that this program 
will generate: 'T' followed by the related room number, 
and a sequential counting number if several doors 
belong to one room !!!
___________________________________________________________
Description:
This script assigns room numbers to doors based on the 
connected rooms' names and creates unique door numbers. 
It also provides a list of generated door numbers and 
identifies any doors that didn't receive a number.
___________________________________________________________
How-to:
Press the button.
___________________________________________________________
Prerequisite:
Please ensure that the room numbers are applied 
before executing the script!!!
___________________________________________________________
"""

import clr
clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
clr.AddReference('RevitNodes')
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)
clr.AddReference('System.Drawing')
from System.Drawing import Color, ColorTranslator

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit import DB, UI
from Autodesk.Revit.DB import FilteredElementCollector as FEC
from Autodesk.Revit.DB import Architecture as AR
from Autodesk.Revit.UI import Selection as SEL
from System import *


doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

room_list = [
    'Bad',
    'Dusche/Toilette',
    'Dusche',
    'Toilette',
    'WC',
    'WC Damen',
    'WC Herren',
    'Waschküche',
    'Waschraum',
    'Kind',
    'Eltern',
    'Schlafen',
    'HWR',
    'Wohnen',
    'Essen',
    'Wohnen/Schlafen',
    'Wohnen/Essen',
    'Arbeiten',
    'Abstellraum',
    'Garderobe',
    'Gast',
    'Hausanschlussraum',
    'Hauswirtschaft',
    'Keller',
    'Kochen/Essen',
    'Wohnbereich',
    'Wohndiele',
    'Wohnküche',
    'Zimmer',
    'Ambulanz',
    'Arztzimmer',
    'Bewohnbarer Raum',
    'Ausstellung',
    'Behandlung',
    'Bereitschaftsraum',
    'Diagnostik',
    'Bühne',
    'Cafeteria',
    'Dienstzimmer',
    'Entwicklung',
    'Erstversorgung',
    'Gemeinschaftsraum',
    'Heizung',
    'Kiosk',
    'Krankenzimmer',
    'Labor',
    'Mehrzweck',
    'Operationssaal',
    'Pausenraum',
    'Pflege',
    'Praxis',
    'Rehabilitation',
    'Röntgen',
    'Ruheraum',
    'Sanitär',
    'Sprechzimmer',
    'Schwesternzimmer',
    'Therapie',
    'Warteraum',
    'Untersuchung',
    'Zählerraum',
    'Zuschauerraum',
    'Großraumbüro',
    'Büro',
    'Aula',
    'Besprechung',
    'Bibliothek',
    'EDV Anlagen',
    'Essbereich',
    'Forum',
    'Kassenraum',
    'Unterricht',
    'Werkhalle',
    'Werkstatt',
    'Hörsaal',
    'Kantine',
    'Klassenraum',
    'Konferenz',
    'Laden',
    'Lehrsaal',
    'Messehalle',
    'Sekretariat',
    'Seminarraum',
    'Sitzungssaal',
    'Sporthalle',
    'Produktion',
    'Nasszelle',
    'Küche',
    'Speiseausgabe',
    'Speisekammer',
    'Speiseraum',
    'Teeküche',
    'Sauna',
    'Schlafbereich',
    'Schwimmbad',
    'Umkleide',
    'Wickelraum',
    'Wellnessbereich',
    'Windfang',
    'Wintergarten',
    'Aufzug',
    'Aufzugsschacht',
    'Fahrstuhl',
    'Fahrradraum',
    'Lager',
    'Techn. Anlagen',
    'Technik',
    'Fernmeldetechnik',
    'HWR',
    'PuMi',
    'Technikraum',
    'Te. Wärmepumpe',
    'Versammlung',
    'Lufttechnik',
    'BMA/BOS',
    'Müll',
    'Müllraum',
    'HA/TW',
    'HA/ELT',
    'TR.-Empfang',
    'Verkauf',
    'Versand',
    'Vorplatz',
    'Vorrat',
    'Vorraum',
    'Vorzimmer',
    'Balkon',
    'Carport',
    'Dachgarten',
    'Dachraum',
    'Doppelgarage',
    'Durchfahrt',
    'Loggia',
    'Rampe',
    'Schleuse',
    'Terrasse',
    'Fluchtbalkon',
    'Garage',
    'Freisitz',
    'Zufahrt',
    # 'Treppenhaus',
    'Treppenraum',
    'Eingang',
    'Gang',
    'Hausflur',
    'Flur',
    'Foyer',
    'Galerie',
    'Diele',
    'Korridor'
]

room_list.reverse()
# print(room_list)
doors = [door for door in FEC(doc).OfCategory(
    DB.BuiltInCategory.OST_Doors).WhereElementIsNotElementType().ToElements() if door.Parameter[
    DB.BuiltInParameter.PHASE_CREATED].AsValueString() == 'Neu']

phase_created = [door.Parameter[DB.BuiltInParameter.PHASE_CREATED].AsElementId() for door in doors]
door_phase = doc.GetElement(phase_created[0])

door_name = []
repeated_door_names = []
repeated_doors = []
room_name_check = []
temp_numbers = []
doors_not_named = []

with DB.Transaction(doc, 'Assign Door Number') as t:
    t.Start()
    for door in doors:
        number_parameter = door.LookupParameter('H_TÜ_Türnummer')
        if door.ToRoom[door_phase] != None and door.FromRoom[door_phase] != None:
            to_room = door.ToRoom[door_phase].Parameter[DB.BuiltInParameter.ROOM_NAME].AsString()
            from_room = door.FromRoom[door_phase].Parameter[DB.BuiltInParameter.ROOM_NAME].AsString()
            for name in room_list:
                if name in to_room or name in from_room:
                    if name in to_room:
                        door_number_room = door.ToRoom[door_phase].LookupParameter('H_RA_Raumnummer').AsString()[1:]
                    else:
                        door_number_room = door.FromRoom[door_phase].LookupParameter('H_RA_Raumnummer').AsString()[1:]
                    number_parameter.Set('T' + door_number_room)
                    continue
                else:
                    pass
        elif door.ToRoom[door_phase] == None and door.FromRoom[door_phase] != None:
            from_room = door.FromRoom[door_phase].Parameter[DB.BuiltInParameter.ROOM_NAME].AsString()
            for name in room_list:
                if name in from_room:
                    door_number_room = door.FromRoom[door_phase].LookupParameter('H_RA_Raumnummer').AsString()[1:]
                    number_parameter.Set('T' + door_number_room)
                    continue
                else:
                    pass
        elif door.FromRoom[door_phase] == None and door.ToRoom[door_phase] != None:
            to_room = door.ToRoom[door_phase].Parameter[DB.BuiltInParameter.ROOM_NAME].AsString()
            for name in room_list:
                if name in to_room:
                    door_number_room = door.ToRoom[door_phase].LookupParameter('H_RA_Raumnummer').AsString()[1:]
                    number_parameter.Set('T' + door_number_room)
                    continue
                else:
                    pass
        else:
            doors_not_named.append(door)
        if door.LookupParameter('H_TÜ_Türnummer').AsString() not in door_name and door.LookupParameter(
            'H_TÜ_Türnummer').AsString() != None:
            door_name.append(door.LookupParameter('H_TÜ_Türnummer').AsString())
        else:
            repeated_doors.append(door)
            repeated_door_names.append(door.LookupParameter('H_TÜ_Türnummer').AsString())
    for door, [i, value] in zip(repeated_doors, enumerate(repeated_door_names)):
        if value != None:
            totalcount = repeated_door_names.count(value)
            count = repeated_door_names[:i].count(value)
            door_number = value + '.' + str(count+1) if totalcount > 0 else value
            number_parameter = door.LookupParameter('H_TÜ_Türnummer')
            number_parameter.Set(door_number)
            door_name.append(door_number)
        elif door.ToRoom[door_phase] != None and door.FromRoom[door_phase]!= None:
            room_name_check.append([door.ToRoom[door_phase].Id.IntegerValue, 
                                    door.FromRoom[door_phase].Id.IntegerValue])
        else:
            pass
    t.Commit()


print('Following door numbers were generated:')
for name in door_name:
    print(name)

if len(doors_not_named) > 0:
    print('***')
    print('The following doors didn`t gen a number:')
    for door in doors_not_named:
        print(door.Id)
    print('Please check a "Türliste"')
else:
    pass
    # print('Repeated names:')  
    # for name in repeated_door_names:
    #     print(name)





