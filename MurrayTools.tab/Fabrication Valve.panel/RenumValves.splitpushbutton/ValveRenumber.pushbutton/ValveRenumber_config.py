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

    # Read the previous input from file
    try:
        with open(filepath, 'r') as f:
            PrevInput = f.read().strip()
    except:
        PrevInput = "1"

    # Function to extract prefix and numeric part from a string
    def extract_number_and_prefix(val):
        import re
        # Find the last sequence of digits in the string
        match = re.search(r'(\d+)$', val)
        if match:
            num = match.group(1)
            prefix = val[:match.start()]
            return prefix, int(num), len(num)
        else:
            # If no number found, treat the whole string as prefix and start with 1
            return val, 0, 0

    # Get initial value by incrementing PrevInput
    prefix, num, num_length = extract_number_and_prefix(PrevInput)
    initial_value = "{}{}".format(prefix, str(num + 1).zfill(num_length or 1))

    # Ask user for input, defaulting to incremented value
    value = forms.ask_for_string(default=initial_value, prompt='Enter Valve Number:', title='Valve Number')

    # Write the user-provided value back to the file
    with open(filepath, 'w') as f:
        f.write(value)

    def distance_between_parts(part1, part2):
        point1 = part1.Origin
        point2 = part2.Origin
        return point1.DistanceTo(point2)

    valves_to_renumber_sorted = sorted(valves_to_renumber, key=lambda x: distance_between_parts(selected_valve, x))

    # Extract prefix and starting number from the user-provided value
    prefix, start_num, num_length = extract_number_and_prefix(value)
    num_length = num_length or 1  # Default to 1 if no number was found

    # Increment logic for renumbering
    numincrement = start_num - 1
    for valve in valves_to_renumber_sorted:
        ST = valve.ServiceType
        AL = valve.Alias
        if ST == 53 and AL not in ['STRAINER', 'CHECK', 'BALANCE']:
            # Increment the number
            numincrement += 1
            # Format the new valve number with leading zeros
            newvalvenumber = "{}{}".format(prefix, str(numincrement).zfill(num_length))
            
            # Set the parameters
            set_parameter_by_name(valve, 'FP_Valve Number', newvalvenumber)
            set_parameter_by_name(valve, 'Mark', newvalvenumber)
            set_customdata_by_custid(valve, 2, newvalvenumber)

    # Write the last assigned number back to the file
    try:
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