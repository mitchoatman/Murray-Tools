import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import Window, Thickness, WindowStyle, WindowStartupLocation, GridLength, ResizeMode
from System.Windows.Controls import Label, ListBox, Grid, RowDefinition
from System.Windows.Media import Brushes
from System.Collections.Generic import List
from System import Action
import System.Windows.Threading

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, ElementId
from Autodesk.Revit.UI import TaskDialog

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView


def convert_fractions(string):
    string = string.replace('"', '')
    tokens = string.split()
    integer_part = 0
    fractional_part = 0.0

    for token in tokens:
        if " " in token:
            integer_part_str, fractional_part_str = token.split(" ")
            integer_part += int(integer_part_str)
            fractional_part_str = fractional_part_str.replace('/', '')
            fractional_part += float(fractional_part_str)
        elif "/" in token:
            numerator, denominator = token.split("/")
            fractional_part += float(numerator) / float(denominator)
        else:
            integer_part += float(token)

    return integer_part + fractional_part


def get_parameter_value_by_name(element, parameterName):
    return element.LookupParameter(parameterName).AsString()


def get_parameter_value_by_name_AsValueString(element, parameterName):
    param = element.LookupParameter(parameterName)
    if param and param.HasValue:
        return param.AsValueString() or param.AsString()
    return ""


hangers = (
    FilteredElementCollector(doc, curview.Id)
    .OfCategory(BuiltInCategory.OST_FabricationHangers)
    .WhereElementIsNotElementType()
    .ToElements()
)

hanger_data = []

for hanger in hangers:
    hosted_info = hanger.GetHostedInfo()
    if hosted_info is None or hosted_info.HostId == ElementId.InvalidElementId:
        family_name = get_parameter_value_by_name_AsValueString(hanger, 'Family')
        if 'Trapeze' not in family_name:
            hanger_data.append(("{}: {}".format(family_name, hanger.Id), hanger.Id))


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
        self.element_map = {text: eid for text, eid in hanger_data}

        grid = Grid()
        grid.Margin = Thickness(10)
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(2)))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
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
        for text, _ in hanger_data:
            self.listbox.Items.Add(text)
        self.listbox.MouseDoubleClick += self.select_element
        Grid.SetRow(self.listbox, 2)
        grid.Children.Add(self.listbox)

    def select_element(self, sender, args):
        selected_text = self.listbox.SelectedItem
        if not selected_text:
            return

        eid = self.element_map[selected_text]
        element = self.doc.GetElement(eid)

        if element:
            self.uidoc.Selection.SetElementIds(List[ElementId]([eid]))
            self.uidoc.ShowElements(eid)
            self.Close()
        else:
            TaskDialog.Show("Error", "Element not found.")


if hanger_data:
    form = HangerListForm(hanger_data, doc, uidoc)
    form.Show()
    while form.IsVisible:
        dispatcher = System.Windows.Threading.Dispatcher.CurrentDispatcher
        dispatcher.Invoke(
            System.Windows.Threading.DispatcherPriority.Background,
            Action(lambda: None)
        )
else:
    TaskDialog.Show("No Matches", "No hangers without a host found.")