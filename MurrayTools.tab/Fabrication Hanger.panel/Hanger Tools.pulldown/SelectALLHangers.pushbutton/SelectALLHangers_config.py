import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
from System.Windows import Application, Window, Thickness, WindowStyle, WindowStartupLocation, GridLength, ResizeMode
from System.Windows.Controls import Label, ListBox, Grid, RowDefinition
from System.Windows.Media import Brushes
from System.Collections.Generic import List
from System import Action
import System.Windows.Threading

from Autodesk.Revit import DB
from Autodesk.Revit.DB import FabricationPart, FabricationAncillaryUsage, Transaction, TransactionGroup, FilteredElementCollector, ElementCategoryFilter, BuiltInCategory, ElementId
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import *

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

def convert_fractions(string):
    # Remove double quotes (") from the input string.
    string = string.replace('"', '')
    # Split the string into a list of tokens.
    tokens = string.split()
    # Initialize variables to keep track of the integer and fractional parts.
    integer_part = 0
    fractional_part = 0.0
    # Iterate over the tokens and convert the mixed number to a float.
    for token in tokens:
        if " " in token:
            # Split the mixed number into integer and fractional parts.
            integer_part_str, fractional_part_str = token.split(" ")
            # Convert the integer part to an integer.
            integer_part += int(integer_part_str)
            # Convert the fractional part to a float.
            fractional_part_str = fractional_part_str.replace('/', '')
            fractional_part += float(fractional_part_str)
        elif "/" in token:
            # If the token is just a fraction, convert it to a float and add it to the fractional part.
            numerator, denominator = token.split("/")
            fractional_part += float(numerator) / float(denominator)
        else:
            # If the token is a standalone number, add it to the integer part.
            integer_part += float(token)
    # Calculate the final result by adding the integer and fractional parts together.
    result = integer_part + fractional_part
    return result

def get_parameter_value_by_name(element, parameterName):
    return element.LookupParameter(parameterName).AsString()

def get_parameter_value_by_name_AsValueString(element, parameterName):
    param = element.LookupParameter(parameterName)
    if param and param.HasValue:
        return param.AsValueString() or param.AsString()
    return ""

# Creating collector instance and collecting all the fabrication hangers from the model
hangers = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationHangers) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

# Collect hanger data for those without a host
hanger_data = []
hangers_without_host = False
for hanger in hangers:
    hosted_info = hanger.GetHostedInfo()
    if hosted_info is None or hosted_info.HostId == DB.ElementId.InvalidElementId:
        hangers_without_host = True
        family_name = get_parameter_value_by_name_AsValueString(hanger, 'Family')
        text = "{}: {}".format(family_name, hanger.Id)
        hanger_data.append((text, hanger.Id))

# WPF Window
class HangerListForm(Window):
    def __init__(self, hanger_data, doc, uidoc):
        self.Title = "Hangers Without Host"
        self.Width = 400
        self.Height = 400
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.ResizeMode = ResizeMode.NoResize
        self.Topmost = True

        self.doc = doc
        self.uidoc = uidoc

        grid = Grid()
        grid.Margin = Thickness(10)

        # Define layout: label + tight spacer + listbox
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))  # Label row
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(2)))    # Spacer
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))  # ListBox row

        self.Content = grid

        label = Label()
        label.Content = "Double Click a hanger to zoom to it:"
        label.Margin = Thickness(0)
        label.Foreground = Brushes.Black
        Grid.SetRow(label, 0)
        grid.Children.Add(label)

        self.listbox = ListBox()
        self.listbox.Margin = Thickness(0)
        self.listbox.Height = 300
        for (text, eid) in hanger_data:
            self.listbox.Items.Add(text)
        Grid.SetRow(self.listbox, 2)
        grid.Children.Add(self.listbox)

        # Map text to element id
        self.element_map = {text: eid for (text, eid) in hanger_data}
        self.listbox.MouseDoubleClick += self.select_element

    def select_element(self, sender, args):
        selected_text = self.listbox.SelectedItem
        if selected_text:
            eid = self.element_map[selected_text]
            element = self.doc.GetElement(eid)
            if element:
                self.uidoc.Selection.SetElementIds(List[ElementId]([eid]))
                self.uidoc.ShowElements(eid)
                self.Close()
            else:
                TaskDialog.Show("Error", "Element not found.")

# Launch UI
if hanger_data:
    form = HangerListForm(hanger_data, doc, uidoc)
    form.Show()
    while form.IsVisible:
        Dispatcher = System.Windows.Threading.Dispatcher.CurrentDispatcher
        Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Background, Action(lambda: None))
else:
    TaskDialog.Show("No Matches", "No hangers without a host found.")