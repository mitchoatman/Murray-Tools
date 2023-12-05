import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter, FilterStringLessOrEqual, FilterStringRule, \
ParameterValueProvider, ElementId, FilterStringBeginsWith, Transaction, FilterStringEquals, \
ElementParameterFilter, ParameterValueProvider, LogicalOrFilter, TransactionGroup, FabricationPart
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from pyrevit import revit, DB, forms
from SharedParam.Add_Parameters import Shared_Params
import os
from Parameters.Get_Set_Params import get_parameter_value_by_name_AsString
Shared_Params()

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
fec = FilteredElementCollector
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float (RevitVersion)

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_ValveRENumber.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)
if not os.path.exists(filepath):
    f = open((filepath), 'w')
    f.write('1')
    f.close()

f = open((filepath), 'r')
PrevInput = f.read()
f.close()

pipe_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationPipework) \
                   .WhereElementIsNotElementType() \
                   .ToElements()
SrvcList = list()

for Item in pipe_collector:
    servicename = get_parameter_value_by_name_AsString(Item, 'Fabrication Service Name')
    SrvcList.append(servicename)

servicelist = forms.SelectFromList.show(set(SrvcList), multiselect=True, button_name='Isolate Service(s)')

list_of_filters = list()

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

if servicelist:
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

        t = Transaction(doc, "Isolate Services")
        t.Start()

        TempIsolate = curview.IsolateElementsTemporary(analyticalCollector)

        t.Commit()

    pipe_collectorv = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationPipework) \
                       .WhereElementIsNotElementType() \
                       .ToElements()

    #FUNCTION TO SET PARAMETER VALUE
    def set_parameter_by_name(element, parameterName, value):
        element.LookupParameter(parameterName).Set(value)
        
    def set_customdata_by_custid(fabpart, custid, value):
        fabpart.SetPartCustomDataText(custid, value)

    def renumber_valves_by_proximity(selected_valve):
        collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationPipework)
        valves_to_renumber = collector.ToElements()

        #This displays dialog
        value = forms.ask_for_string(default=PrevInput, prompt='Enter Valve Number:', title='Valve Number')

        f = open((filepath), 'w')
        f.write(value)
        f.close()

        def distance_between_parts(part1, part2):
            point1 = part1.Origin
            point2 = part2.Origin
            return point1.DistanceTo(point2)

        valves_to_renumber_sorted = sorted(valves_to_renumber, key=lambda x: distance_between_parts(selected_valve, x))
        
         
        
        if "-" in value:
            #splits the spoolname
            valuesplit = value.rsplit('-', 1)
            
            #gets the length of characters for number
            valvenumlength = (len(valuesplit[-1]))

            #gets the first half of spool name
            firstpart = valuesplit[0]
            
            #converts number from string to integer
            valuenum = int(float(valuesplit[-1]))
            
            numincrement = valuenum - 1
            
            for valve in valves_to_renumber_sorted:
                ST = valve.ServiceType
                if ST == 53:
                    #increments valve number  by 1
                    numincrement = numincrement + 1

                    #print numincrement
                    #converts valve number back into string and fills in leading zeros
                    lastpart = str(numincrement).zfill(valvenumlength)

                    #combines both halfs of valve number
                    newvalvenumber = firstpart + "-" + lastpart
                        
                    set_parameter_by_name(valve, 'FP_Valve Number', newvalvenumber)
                    set_parameter_by_name(valve, 'Mark', newvalvenumber)
                    set_customdata_by_custid(valve, 2, newvalvenumber)
        else:
            #converts number from string to integer
            valuenum = int(float(value[-1]))

            #gets the length of characters for number
            valvenumlength = (len(value))

            numincrement = valuenum - 1
            
            for valve in valves_to_renumber_sorted:
                ST = valve.ServiceType
                if ST == 53:
                    #increments valve number  by 1
                    numincrement = numincrement + 1

                    #print numincrement
                    #converts valve number back into string and fills in leading zeros
                    lastpart = str(numincrement).zfill(valvenumlength)

                    #combines both halfs of valve number
                    newvalvenumber = lastpart
                        
                    set_parameter_by_name(valve, 'FP_Valve Number', newvalvenumber)
                    set_parameter_by_name(valve, 'Mark', newvalvenumber)
                    set_customdata_by_custid(valve, 2, newvalvenumber)

        f = open((filepath), 'w')
        f.write(newvalvenumber)
        f.close()

    t = Transaction(doc, 'Set Valve Number')
    #Start Transaction
    t.Start()

    selected_part_ref = uidoc.Selection.PickObject(ObjectType.Element, "Select Valve to start numbering from")
    selected_valve = doc.GetElement(selected_part_ref)
    renumber_valves_by_proximity(selected_valve)

    curview.DisableTemporaryViewMode(DB.TemporaryViewMode.TemporaryHideIsolate)

    #End Transaction
    t.Commit()
