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

def set_customdata_by_custid(fabpart, custid, value):
    fabpart.SetPartCustomDataText(custid, value)

def create_filter_2023_newer(key_parameter, element_value):
    """Function to create a filter from builtinParameter and Value."""
    f_parameter = ParameterValueProvider(ElementId(key_parameter))
    f_parameter_value = element_value
    f_rule = FilterStringRule(f_parameter, FilterStringEquals(), f_parameter_value)
    my_filter = ElementParameterFilter(f_rule)
    return my_filter

def create_filter_2022_older(key_parameter, element_value):
    """Function to create a filter from builtinParameter and Value."""
    f_parameter = ParameterValueProvider(ElementId(key_parameter))
    f_parameter_value = element_value
    caseSensitive = False
    f_rule = FilterStringRule(f_parameter, FilterStringEquals(), f_parameter_value, caseSensitive)
    my_filter = ElementParameterFilter(f_rule)
    return my_filter

pipe_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationPipework) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

SrvcList = list()

for Item in pipe_collector:
    servicename = get_parameter_value_by_name_AsString(Item, 'Fabrication Service Name')
    SrvcList.append(servicename)

servicelist = forms.SelectFromList.show(set(SrvcList), multiselect=True, button_name='Select Service(s) Drawn in SS')

list_of_filters = list()

if RevitINT > 2022:
    for fp_servicename in servicelist:
        cat_filter = create_filter_2023_newer(key_parameter = BuiltInParameter.FABRICATION_SERVICE_NAME, element_value = str(fp_servicename))
        list_of_filters.Add(cat_filter)
else:
    for fp_servicename in servicelist:
        cat_filter = create_filter_2022_older(key_parameter = BuiltInParameter.FABRICATION_SERVICE_NAME, element_value = str(fp_servicename))
        list_of_filters.Add(cat_filter)

if list_of_filters:
    multiple_filters = LogicalOrFilter(list_of_filters)

    analyticalCollector = FilteredElementCollector(doc).WherePasses(multiple_filters).ToElementIds()

tg = TransactionGroup(doc, "Quick CS Pointload")
tg.Start()

t = Transaction(doc, "Isolate Services")
t.Start()

curview.IsolateElementsTemporary(analyticalCollector)

t.Commit()
    
hanger_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationHangers) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

t = Transaction(doc, 'Set CU Pointload Values')
#Start Transaction
t.Start()

for hanger in hanger_collector:
    hosted_info = hanger.GetHostedInfo().HostId
    try:
        HangerSize = get_parameter_value_by_name_AsString(doc.GetElement(hosted_info), 'Size')
        if HangerSize == '1/2"':
            set_parameter_by_name(hanger, 'FP_Pointload', 5)
            set_customdata_by_custid(hanger, 7, '5')
        if HangerSize == '3/4"':
            set_parameter_by_name(hanger, 'FP_Pointload', 9)
            set_customdata_by_custid(hanger, 7, '9')
        if HangerSize == '1"':
            set_parameter_by_name(hanger, 'FP_Pointload', 15)
            set_customdata_by_custid(hanger, 7, '15')
        if HangerSize == '1 1/4"':
            set_parameter_by_name(hanger, 'FP_Pointload', 26)
            set_customdata_by_custid(hanger, 7, '26')
        if HangerSize == '1 1/2"':
            set_parameter_by_name(hanger, 'FP_Pointload', 31)
            set_customdata_by_custid(hanger, 7, '31')
        if HangerSize == '2"':
            set_parameter_by_name(hanger, 'FP_Pointload', 43)
            set_customdata_by_custid(hanger, 7, '43')
        if HangerSize == '2 1/2"':
            set_parameter_by_name(hanger, 'FP_Pointload', 59)
            set_customdata_by_custid(hanger, 7, '59')
        if HangerSize == '3"':
            set_parameter_by_name(hanger, 'FP_Pointload', 80)
            set_customdata_by_custid(hanger, 7, '80')
        if HangerSize == '4"':
            set_parameter_by_name(hanger, 'FP_Pointload', 12)
            set_customdata_by_custid(hanger, 7, '12')
        if HangerSize == '6"':
            set_parameter_by_name(hanger, 'FP_Pointload', 24)
            set_customdata_by_custid(hanger, 7, '24')
        if HangerSize == '8"':
            set_parameter_by_name(hanger, 'FP_Pointload', 37)
            set_customdata_by_custid(hanger, 7, '37')
    except:
        output = script.get_output()
        print('{}: {}'.format('Disconnected Hanger', output.linkify(hanger.Id)))

#End Transaction
t.Commit()

t = Transaction(doc, "Reset View")
t.Start()

curview.DisableTemporaryViewMode(DB.TemporaryViewMode.TemporaryHideIsolate)

t.Commit()
#End Transaction Group
tg.Assimilate()