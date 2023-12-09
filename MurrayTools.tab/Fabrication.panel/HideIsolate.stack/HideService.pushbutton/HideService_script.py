import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter, \
ParameterValueProvider, ElementId, FilterStringBeginsWith, Transaction, FilterStringEquals, \
FilterStringLessOrEqual, FilterStringRule, ElementParameterFilter, ParameterValueProvider, LogicalOrFilter
from pyrevit import forms
from Parameters.Get_Set_Params import get_parameter_value_by_name_AsString

doc = __revit__.ActiveUIDocument.Document
DB = Autodesk.Revit.DB
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float (RevitVersion)

pipe_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationPipework) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

duct_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationDuctwork) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

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

SrvcList = list()

for Item in pipe_collector:
    servicename = get_parameter_value_by_name_AsString(Item, 'Fabrication Service Name')
    SrvcList.append(servicename)

for Item in duct_collector:
    servicename = get_parameter_value_by_name_AsString(Item, 'Fabrication Service Name')
    SrvcList.append(servicename)

servicelist = forms.SelectFromList.show(set(SrvcList), title="Choose Service(s)", multiselect=True, button_name='Hide Service(s)')

list_of_filters = list()
try:
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

        t = Transaction(doc, "Hide Services")
        t.Start()

        curview.HideElementsTemporary(analyticalCollector)

        t.Commit()
except:
    pass