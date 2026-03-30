import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter, FilterStringLessOrEqual, FilterStringRule, \
ParameterValueProvider, ElementId, FilterStringBeginsWith, Transaction, FilterStringEquals, \
ElementParameterFilter, ParameterValueProvider, LogicalOrFilter, TransactionGroup, FabricationPart
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from pyrevit import revit, DB, forms
from Parameters.Add_SharedParameters import Shared_Params
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
            list_of_filters.append(cat_filter)  # Changed Add to append
    else:
        for fp_servicename in servicelist:
            cat_filter = create_filter_2022_older(key_parameter = BuiltInParameter.FABRICATION_SERVICE_NAME, element_value = str(fp_servicename))
            list_of_filters.append(cat_filter)  # Changed Add to append

    if list_of_filters:
        multiple_filters = LogicalOrFilter(list_of_filters)

        analyticalCollector = FilteredElementCollector(doc).WherePasses(multiple_filters).ToElementIds()

        t = Transaction(doc, "Isolate Services")
        t.Start()

        TempIsolate = curview.IsolateElementsTemporary(analyticalCollector)

        t.Commit()

    #FUNCTION TO SET PARAMETER VALUE
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

    t = Transaction(doc, 'Set Valve Number')
    #Start Transaction
    t.Start()

    selected_part_ref = uidoc.Selection.PickObject(ObjectType.Element, "Select Valve to start numbering from")
    selected_valve = doc.GetElement(selected_part_ref)
    renumber_valves_by_proximity(selected_valve)

    curview.DisableTemporaryViewMode(DB.TemporaryViewMode.TemporaryHideIsolate)

    #End Transaction
    t.Commit()
