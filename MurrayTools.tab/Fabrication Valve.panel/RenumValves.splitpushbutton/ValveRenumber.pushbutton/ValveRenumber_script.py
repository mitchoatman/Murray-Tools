import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter, FilterStringRule, \
    ParameterValueProvider, ElementId, Transaction, FilterStringEquals, ElementParameterFilter, FabricationPart
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.UI import TaskDialog
from Parameters.Add_SharedParameters import Shared_Params
import os
from Parameters.Get_Set_Params import get_parameter_value_by_name_AsString
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
from System.Windows import Application, Window, Thickness, WindowStyle, ResizeMode, WindowStartupLocation, HorizontalAlignment
from System.Windows.Controls import Label, TextBox, Button, Grid, RowDefinition

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

# WPF dialog class (programmatic, no XAML)
class ValveNumberForm(Window):
    def __init__(self, default_value):
        self.Title = "Valve Number"
        self.Width = 300
        self.Height = 150
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.result = None


        # Grid layout
        grid = Grid()
        grid.Margin = Thickness(10)
        row1 = RowDefinition()
        row2 = RowDefinition()
        row3 = RowDefinition()
        grid.RowDefinitions.Add(row1)
        grid.RowDefinitions.Add(row2)
        grid.RowDefinitions.Add(row3)
        self.Content = grid

        # Label (moved up 5px by setting top margin to -5)
        label = Label()
        label.Content = "Enter Starting Valve Number:"
        label.Margin = Thickness(0, -5, 0, 10)
        Grid.SetRow(label, 0)
        grid.Children.Add(label)

        # TextBox
        self.textbox = TextBox()
        self.textbox.Text = default_value
        self.textbox.Margin = Thickness(0, 0, 0, 10)
        Grid.SetRow(self.textbox, 1)
        grid.Children.Add(self.textbox)

        # OK Button (centered)
        ok_button = Button()
        ok_button.Content = "OK"
        ok_button.Width = 75
        ok_button.HorizontalAlignment = HorizontalAlignment.Center
        ok_button.Click += self.on_ok
        Grid.SetRow(ok_button, 2)
        grid.Children.Add(ok_button)

    def OnContentRendered(self, e):
        Window.OnContentRendered(self, e)
        self.textbox.Focus()
        self.textbox.SelectAll()

    def on_ok(self, sender, args):
        self.result = self.textbox.Text
        self.DialogResult = True
        self.Close()

class ValveSelectionFilter(ISelectionFilter):
    def AllowElement(self, e):
        if not isinstance(e, FabricationPart):
            return False
        ST = e.ServiceType
        AL = e.Alias
        return ST == 53 and AL not in ['STRAINER', 'CHECK', 'BALANCE', 'PLUG']

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
        match = re.search(r'(\d+)$', val)
        if match:
            num = match.group(1)
            prefix = val[:match.start()]
            return prefix, int(num), len(num)
        else:
            return val, 0, 0

    # Get initial value by incrementing PrevInput
    prefix, num, num_length = extract_number_and_prefix(PrevInput)
    initial_value = "{}{}".format(prefix, str(num + 1).zfill(num_length or 1))

    # Show WPF dialog (no XAML)
    form = ValveNumberForm(initial_value)
    if form.ShowDialog() and form.DialogResult:
        value = form.result
    else:
        return  # Exit if dialog is cancelled

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
        if ST == 53 and AL not in ['STRAINER', 'CHECK', 'BALANCE', 'PLUG']:
            # Increment the number
            numincrement += 1
            # Format the new valve number with leading zeros
            newvalvenumber = "{}{}".format(prefix, str(numincrement).zfill(num_length))
            
            # Set the parameters
            set_parameter_by_name(valve, 'FP_Valve Number', newvalvenumber)
            set_customdata_by_custid(valve, 2, newvalvenumber)

    # Write the last assigned number back to the file
    try:
        with open(filepath, 'w') as f:
            f.write(newvalvenumber)
    except NameError:
        TaskDialog.Show("No Valves Found", "No Valves Found in View")

# Start main logic
t = Transaction(doc, 'Set Valve Number and Isolate Service')
t.Start()

from Autodesk.Revit.Exceptions import OperationCanceledException

try:
    # Prompt user to select a valve
    selected_part_ref = uidoc.Selection.PickObject(ObjectType.Element, ValveSelectionFilter(), "Select Valve to start numbering from")
    selected_valve = doc.GetElement(selected_part_ref)
except OperationCanceledException:
    t.RollBack()
    TaskDialog.Show("Cancelled", "Operation cancelled by user.")
    import sys
    sys.exit()  # Exit script safely
except Exception as e:
    t.RollBack()
    TaskDialog.Show("Error", str(e))
    import sys
    sys.exit()  # Exit script safely


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