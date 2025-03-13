import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter, FilterStringRule, \
    ParameterValueProvider, ElementId, Transaction, FilterStringEquals, ElementParameterFilter, FabricationPart
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import revit, DB, forms
from SharedParam.Add_Parameters import Shared_Params
import os
from Parameters.Get_Set_Params import get_parameter_value_by_name_AsString

Shared_Params()

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_ValveRENumber.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)
if not os.path.exists(filepath):
    with open(filepath, 'w') as f:
        f.write('1')

with open(filepath, 'r') as f:
    PrevInput = f.read()

# Functions for creating filters based on Revit version
def create_filter_2023_newer(key_parameter, element_value):
    """Function to create a filter from builtinParameter and Value."""
    f_parameter = ParameterValueProvider(ElementId(key_parameter))
    f_rule = FilterStringRule(f_parameter, FilterStringEquals(), element_value)
    return ElementParameterFilter(f_rule)

def create_filter_2022_older(key_parameter, element_value):
    """Function to create a filter from builtinParameter and Value."""
    f_parameter = ParameterValueProvider(ElementId(key_parameter))
    caseSensitive = False
    f_rule = FilterStringRule(f_parameter, FilterStringEquals(), element_value, caseSensitive)
    return ElementParameterFilter(f_rule)

# Function to set parameter values
def set_parameter_by_name(element, parameterName, value):
    element.LookupParameter(parameterName).Set(value)

def set_customdata_by_custid(fabpart, custid, value):
    fabpart.SetPartCustomDataText(custid, value)

def renumber_valves_by_proximity(selected_valve):
    collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationPipework)
    valves_to_renumber = collector.ToElements()

    # Convert PrevInput to integer and increment by 1
    try:
        if "-" in PrevInput:
            valuesplit = PrevInput.rsplit('-', 1)
            start_value = str(int(valuesplit[-1]) + 1)
            initial_value = valuesplit[0] + "-" + start_value.zfill(len(valuesplit[-1]))
        else:
            start_value = str(int(PrevInput) + 1)
            initial_value = start_value.zfill(len(PrevInput))
    except ValueError:
        initial_value = "1"

    # Prompt user for valve number
    value = forms.ask_for_string(default=initial_value, prompt='Enter Valve Number:', title='Valve Number')

    with open(filepath, 'w') as f:
        f.write(value)

    def distance_between_parts(part1, part2):
        point1 = part1.Origin
        point2 = part2.Origin
        return point1.DistanceTo(point2)

    valves_to_renumber_sorted = sorted(valves_to_renumber, key=lambda x: distance_between_parts(selected_valve, x))

    if "-" in value:
        valuesplit = value.rsplit('-', 1)
        valvenumlength = len(valuesplit[-1])
        firstpart = valuesplit[0]
        valuenum = int(float(valuesplit[-1]))
        numincrement = valuenum - 1
        
        for valve in valves_to_renumber_sorted:
            ST = valve.ServiceType
            AL = valve.Alias
            if ST == 53 and AL != 'STRAINER' and AL != 'CHECK' and AL != 'BALANCE':
                numincrement += 1
                lastpart = str(numincrement).zfill(valvenumlength)
                newvalvenumber = firstpart + "-" + lastpart
                set_parameter_by_name(valve, 'FP_Valve Number', newvalvenumber)
                set_parameter_by_name(valve, 'Mark', newvalvenumber)
                set_customdata_by_custid(valve, 2, newvalvenumber)
    else:
        valvenumlength = len(value)
        valuenum = int(float(value))
        numincrement = valuenum - 1
        
        for valve in valves_to_renumber_sorted:
            ST = valve.ServiceType
            AL = valve.Alias
            if ST == 53 and AL != 'STRAINER' and AL != 'CHECK':
                numincrement += 1
                lastpart = str(numincrement).zfill(valvenumlength)
                newvalvenumber = lastpart
                set_parameter_by_name(valve, 'FP_Valve Number', newvalvenumber)
                set_parameter_by_name(valve, 'Mark', newvalvenumber)
                set_customdata_by_custid(valve, 2, newvalvenumber)

    try:
        newvalvenumber
        with open(filepath, 'w') as f:
            f.write(newvalvenumber)
    except NameError:
        forms.alert('No Valves Found', ok=True, yes=False, no=False, exitscript=False)

# Start main logic
t = Transaction(doc, 'Set Valve Number and Isolate Service')
t.Start()

# Prompt user to select a valve
selected_part_ref = uidoc.Selection.PickObject(ObjectType.Element, "Select Valve to start numbering from")
selected_valve = doc.GetElement(selected_part_ref)

# Get the Fabrication Service Name from the selected valve
service_name = get_parameter_value_by_name_AsString(selected_valve, 'Fabrication Service Name')

if service_name:
    # Create filter based on the selected valve's service name
    if RevitINT > 2022:
        service_filter = create_filter_2023_newer(BuiltInParameter.FABRICATION_SERVICE_NAME, service_name)
    else:
        service_filter = create_filter_2022_older(BuiltInParameter.FABRICATION_SERVICE_NAME, service_name)

    # Isolate elements with the same service name
    analyticalCollector = FilteredElementCollector(doc).WherePasses(service_filter).ToElementIds()
    curview.IsolateElementsTemporary(analyticalCollector)

# Renumber valves starting from the selected valve
renumber_valves_by_proximity(selected_valve)

# Disable temporary isolation
curview.DisableTemporaryViewMode(DB.TemporaryViewMode.TemporaryHideIsolate)

t.Commit()