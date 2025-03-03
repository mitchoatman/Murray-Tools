import Autodesk
import sys
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter, FilterStringLessOrEqual, FilterStringRule, \
ParameterValueProvider, ElementId, FilterStringBeginsWith, Transaction, FilterStringEquals, \
ElementParameterFilter, ParameterValueProvider, LogicalOrFilter, TransactionGroup, FabricationPart, FabricationConfiguration
from pyrevit import revit, DB, script, forms
from SharedParam.Add_Parameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString

Shared_Params()

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
fec = FilteredElementCollector
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float (RevitVersion)
Config = FabricationConfiguration.GetFabricationConfiguration(doc)

#This writes to fab part custom data field
def set_customdata_by_custid(fabpart, custid, value):
	fabpart.SetPartCustomDataText(custid, value)

def get_parameter_value_by_name_AsValueString(element, parameterName):
    return element.LookupParameter(parameterName).AsValueString()

hanger_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationHangers) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

t = Transaction(doc, 'Set Pointload Values')
#Start Transaction
t.Start()

for hanger in hanger_collector:
    hosted_info = hanger.GetHostedInfo().HostId
    try:
        Hostmat = doc.GetElement(hosted_info).Parameter[BuiltInParameter.FABRICATION_PART_MATERIAL].AsValueString()  #Copper: Hard Copper  #Cast Iron: Cast Iron
        if Hostmat == 'Cast Iron: Cast Iron':
            HostSize = get_parameter_value_by_name_AsString(doc.GetElement(hosted_info), 'Size')
            if HostSize == '2"':
                set_parameter_by_name(hanger, 'FP_Pointload', 3)
                set_customdata_by_custid(hanger, 7, '3')
            if HostSize == '3"':
                set_parameter_by_name(hanger, 'FP_Pointload', 5)
                set_customdata_by_custid(hanger, 7, '5')
            if HostSize == '4"':
                set_parameter_by_name(hanger, 'FP_Pointload', 7)
                set_customdata_by_custid(hanger, 7, '7')
            if HostSize == '5"':
                set_parameter_by_name(hanger, 'FP_Pointload', 9)
                set_customdata_by_custid(hanger, 7, '9')
            if HostSize == '6"':
                set_parameter_by_name(hanger, 'FP_Pointload', 12)
                set_customdata_by_custid(hanger, 7, '12')
            if HostSize == '8"':
                set_parameter_by_name(hanger, 'FP_Pointload', 20)
                set_customdata_by_custid(hanger, 7, '20')
            if HostSize == '10"':
                set_parameter_by_name(hanger, 'FP_Pointload', 30)
                set_customdata_by_custid(hanger, 7, '30')
            if HostSize == '12"':
                set_parameter_by_name(hanger, 'FP_Pointload', 42)
                set_customdata_by_custid(hanger, 7, '42')
            if HostSize == '15"':
                set_parameter_by_name(hanger, 'FP_Pointload', 65)
                set_customdata_by_custid(hanger, 7, '65')

        if Hostmat == 'Copper: Hard Copper':
            HostSize = get_parameter_value_by_name_AsString(doc.GetElement(hosted_info), 'Size')
            if HostSize == '1/2"':
                set_parameter_by_name(hanger, 'FP_Pointload', 1)
                set_customdata_by_custid(hanger, 7, '1')
            if HostSize == '3/4"':
                set_parameter_by_name(hanger, 'FP_Pointload', 1)
                set_customdata_by_custid(hanger, 7, '1')
            if HostSize == '1"':
                set_parameter_by_name(hanger, 'FP_Pointload', 2)
                set_customdata_by_custid(hanger, 7, '2')
            if HostSize == '1 1/4"':
                set_parameter_by_name(hanger, 'FP_Pointload', 2)
                set_customdata_by_custid(hanger, 7, '2')
            if HostSize == '1 1/2"':
                set_parameter_by_name(hanger, 'FP_Pointload', 2)
                set_customdata_by_custid(hanger, 7, '2')
            if HostSize == '2"':
                set_parameter_by_name(hanger, 'FP_Pointload', 5)
                set_customdata_by_custid(hanger, 7, '5')
            if HostSize == '2 1/2"':
                set_parameter_by_name(hanger, 'FP_Pointload', 6)
                set_customdata_by_custid(hanger, 7, '6')
            if HostSize == '3"':
                set_parameter_by_name(hanger, 'FP_Pointload', 8)
                set_customdata_by_custid(hanger, 7, '8')
            if HostSize == '4"':
                set_parameter_by_name(hanger, 'FP_Pointload', 15)
                set_customdata_by_custid(hanger, 7, '15')
            if HostSize == '6"':
                set_parameter_by_name(hanger, 'FP_Pointload', 30)
                set_customdata_by_custid(hanger, 7, '30')
            if HostSize == '8"':
                set_parameter_by_name(hanger, 'FP_Pointload', 50)
                set_customdata_by_custid(hanger, 7, '50')

        if Hostmat == 'Carbon Steel: Carbon Steel':
            HostSize = get_parameter_value_by_name_AsString(doc.GetElement(hosted_info), 'Size')
            if HostSize == '1/2"':
                set_parameter_by_name(hanger, 'FP_Pointload', 1)
                set_customdata_by_custid(hanger, 7, '1')
            if HostSize == '3/4"':
                set_parameter_by_name(hanger, 'FP_Pointload', 2)
                set_customdata_by_custid(hanger, 7, '2')
            if HostSize == '1"':
                set_parameter_by_name(hanger, 'FP_Pointload', 3)
                set_customdata_by_custid(hanger, 7, '3')
            if HostSize == '1 1/4"':
                set_parameter_by_name(hanger, 'FP_Pointload', 4)
                set_customdata_by_custid(hanger, 7, '4')
            if HostSize == '1 1/2"':
                set_parameter_by_name(hanger, 'FP_Pointload', 5)
                set_customdata_by_custid(hanger, 7, '5')
            if HostSize == '2"':
                set_parameter_by_name(hanger, 'FP_Pointload', 6)
                set_customdata_by_custid(hanger, 7, '6')
            if HostSize == '2 1/2"':
                set_parameter_by_name(hanger, 'FP_Pointload', 10)
                set_customdata_by_custid(hanger, 7, '10')
            if HostSize == '3"':
                set_parameter_by_name(hanger, 'FP_Pointload', 12)
                set_customdata_by_custid(hanger, 7, '12')
            if HostSize == '4"':
                set_parameter_by_name(hanger, 'FP_Pointload', 20)
                set_customdata_by_custid(hanger, 7, '20')
            if HostSize == '6"':
                set_parameter_by_name(hanger, 'FP_Pointload', 36)
                set_customdata_by_custid(hanger, 7, '36')
            if HostSize == '8"':
                set_parameter_by_name(hanger, 'FP_Pointload', 50)
                set_customdata_by_custid(hanger, 7, '50')

        if Hostmat in ['Stainless Steel: 304L', 'Stainless Steel: 316L']:
            HostSize = get_parameter_value_by_name_AsString(doc.GetElement(hosted_info), 'Size')
            if HostSize == '1/2"':
                set_parameter_by_name(hanger, 'FP_Pointload', 1)
                set_customdata_by_custid(hanger, 7, '1')
            if HostSize == '3/4"':
                set_parameter_by_name(hanger, 'FP_Pointload', 1)
                set_customdata_by_custid(hanger, 7, '2')
            if HostSize == '1"':
                set_parameter_by_name(hanger, 'FP_Pointload', 2)
                set_customdata_by_custid(hanger, 7, '2')
            if HostSize == '1 1/4"':
                set_parameter_by_name(hanger, 'FP_Pointload', 3)
                set_customdata_by_custid(hanger, 7, '3')
            if HostSize == '1 1/2"':
                set_parameter_by_name(hanger, 'FP_Pointload', 4)
                set_customdata_by_custid(hanger, 7, '4')
            if HostSize == '2"':
                set_parameter_by_name(hanger, 'FP_Pointload', 5)
                set_customdata_by_custid(hanger, 7, '5')
            if HostSize == '2 1/2"':
                set_parameter_by_name(hanger, 'FP_Pointload', 6)
                set_customdata_by_custid(hanger, 7, '6')
            if HostSize == '3"':
                set_parameter_by_name(hanger, 'FP_Pointload', 8)
                set_customdata_by_custid(hanger, 7, '8')
            if HostSize == '4"':
                set_parameter_by_name(hanger, 'FP_Pointload', 12)
                set_customdata_by_custid(hanger, 7, '12')
            if HostSize == '6"':
                set_parameter_by_name(hanger, 'FP_Pointload', 24)
                set_customdata_by_custid(hanger, 7, '24')
            if HostSize == '8"':
                set_parameter_by_name(hanger, 'FP_Pointload', 37)
                set_customdata_by_custid(hanger, 7, '37')
    except:
        output = script.get_output()
        print('{}: {}'.format((get_parameter_value_by_name_AsValueString(hanger, 'Family')), output.linkify(hanger.Id)))

#End Transaction
t.Commit()
