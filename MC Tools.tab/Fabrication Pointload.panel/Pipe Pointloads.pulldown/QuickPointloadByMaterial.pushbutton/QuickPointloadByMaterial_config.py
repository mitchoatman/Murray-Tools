import Autodesk
import sys
import clr

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter, FilterStringLessOrEqual, FilterStringRule, \
ParameterValueProvider, ElementId, FilterStringBeginsWith, Transaction, FilterStringEquals, \
ElementParameterFilter, ParameterValueProvider, LogicalOrFilter, TransactionGroup, FabricationPart, FabricationConfiguration
from pyrevit import revit, DB, script, forms
from Parameters.Add_SharedParameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString

# WPF Imports
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import Window, Thickness, WindowStyle, ResizeMode, WindowStartupLocation, GridLength
from System.Windows.Controls import Label, ListBox, Grid, RowDefinition, Button
from System.Windows.Media import Brushes
import System
from System import Action
import System.Windows.Threading
from System.Collections.Generic import List
from Autodesk.Revit.UI import TaskDialog

Shared_Params()

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
fec = FilteredElementCollector
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)
Config = FabricationConfiguration.GetFabricationConfiguration(doc)

# This writes to fab part custom data field
def set_customdata_by_custid(fabpart, custid, value):
    fabpart.SetPartCustomDataText(custid, value)

def get_parameter_value_by_name_AsValueString(element, parameterName):
    param = element.LookupParameter(parameterName)
    if param and param.HasValue:
        return param.AsValueString() or param.AsString()
    return ""

hanger_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationHangers) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

error_data = []

t = Transaction(doc, 'Set CI Pointload Values')
# Start Transaction
t.Start()

for hanger in hanger_collector:
    family_name = get_parameter_value_by_name_AsValueString(hanger, 'Family')

    # Skip trapeze hangers
    if 'trapeze' in family_name.lower():
        continue

    try:
        hosted_info_obj = hanger.GetHostedInfo()
        if hosted_info_obj is None or hosted_info_obj.HostId == ElementId.InvalidElementId:
            display_text = "{} (ID: {})".format(family_name, hanger.Id)
            error_data.append((display_text, hanger.Id))
            continue

        hosted_info = hosted_info_obj.HostId

        Hostmat = doc.GetElement(hosted_info).Parameter[BuiltInParameter.FABRICATION_PART_MATERIAL].AsValueString()  # Copper: Hard Copper  # Cast Iron: Cast Iron

        if Hostmat == 'Cast Iron: Cast Iron':
            HostSize = get_parameter_value_by_name_AsString(doc.GetElement(hosted_info), 'Size')
            if HostSize == '2"':
                set_parameter_by_name(hanger, 'FP_Pointload', 24.75)
                set_customdata_by_custid(hanger, 7, '3')
            if HostSize == '3"':
                set_parameter_by_name(hanger, 'FP_Pointload', 41.2)
                set_customdata_by_custid(hanger, 7, '5')
            if HostSize == '4"':
                set_parameter_by_name(hanger, 'FP_Pointload', 64.1)
                set_customdata_by_custid(hanger, 7, '7')
            if HostSize == '5"':
                set_parameter_by_name(hanger, 'FP_Pointload', 87.5)
                set_customdata_by_custid(hanger, 7, '9')
            if HostSize == '6"':
                set_parameter_by_name(hanger, 'FP_Pointload', 115.9)
                set_customdata_by_custid(hanger, 7, '12')
            if HostSize == '8"':
                set_parameter_by_name(hanger, 'FP_Pointload', 198.3)
                set_customdata_by_custid(hanger, 7, '20')
            if HostSize == '10"':
                set_parameter_by_name(hanger, 'FP_Pointload', 298)
                set_customdata_by_custid(hanger, 7, '30')
            if HostSize == '12"':
                set_parameter_by_name(hanger, 'FP_Pointload', 420)
                set_customdata_by_custid(hanger, 7, '42')
            if HostSize == '15"':
                set_parameter_by_name(hanger, 'FP_Pointload', 650)
                set_customdata_by_custid(hanger, 7, '65')

        if Hostmat == 'Copper: Hard Copper':
            HostSize = get_parameter_value_by_name_AsString(doc.GetElement(hosted_info), 'Size')
            if HostSize == '1/2"':
                set_parameter_by_name(hanger, 'FP_Pointload', 2.638)
                set_customdata_by_custid(hanger, 7, '1')
            if HostSize == '3/4"':
                set_parameter_by_name(hanger, 'FP_Pointload', 7.56)
                set_customdata_by_custid(hanger, 7, '1')
            if HostSize == '1"':
                set_parameter_by_name(hanger, 'FP_Pointload', 10.68)
                set_customdata_by_custid(hanger, 7, '2')
            if HostSize == '1 1/4"':
                set_parameter_by_name(hanger, 'FP_Pointload', 11.58)
                set_customdata_by_custid(hanger, 7, '2')
            if HostSize == '1 1/2"':
                set_parameter_by_name(hanger, 'FP_Pointload', 16.56)
                set_customdata_by_custid(hanger, 7, '2')
            if HostSize == '2"':
                set_parameter_by_name(hanger, 'FP_Pointload', 42.1)
                set_customdata_by_custid(hanger, 7, '5')
            if HostSize == '2 1/2"':
                set_parameter_by_name(hanger, 'FP_Pointload', 58.8)
                set_customdata_by_custid(hanger, 7, '6')
            if HostSize == '3"':
                set_parameter_by_name(hanger, 'FP_Pointload', 80.3)
                set_customdata_by_custid(hanger, 7, '8')
            if HostSize == '4"':
                set_parameter_by_name(hanger, 'FP_Pointload', 147.5)
                set_customdata_by_custid(hanger, 7, '15')
            if HostSize == '6"':
                set_parameter_by_name(hanger, 'FP_Pointload', 292.8)
                set_customdata_by_custid(hanger, 7, '30')
            if HostSize == '8"':
                set_parameter_by_name(hanger, 'FP_Pointload', 500)
                set_customdata_by_custid(hanger, 7, '50')

        if Hostmat == 'Carbon Steel: Carbon Steel':
            HostSize = get_parameter_value_by_name_AsString(doc.GetElement(hosted_info), 'Size')
            if HostSize == '1/2"':
                set_parameter_by_name(hanger, 'FP_Pointload', 8.7)
                set_customdata_by_custid(hanger, 7, '1')
            if HostSize == '3/4"':
                set_parameter_by_name(hanger, 'FP_Pointload', 14.32)
                set_customdata_by_custid(hanger, 7, '2')
            if HostSize == '1"':
                set_parameter_by_name(hanger, 'FP_Pointload', 21.36)
                set_customdata_by_custid(hanger, 7, '3')
            if HostSize == '1 1/4"':
                set_parameter_by_name(hanger, 'FP_Pointload', 36.2)
                set_customdata_by_custid(hanger, 7, '4')
            if HostSize == '1 1/2"':
                set_parameter_by_name(hanger, 'FP_Pointload', 42.4)
                set_customdata_by_custid(hanger, 7, '5')
            if HostSize == '2"':
                set_parameter_by_name(hanger, 'FP_Pointload', 58.9)
                set_customdata_by_custid(hanger, 7, '6')
            if HostSize == '2 1/2"':
                set_parameter_by_name(hanger, 'FP_Pointload', 91)
                set_customdata_by_custid(hanger, 7, '10')
            if HostSize == '3"':
                set_parameter_by_name(hanger, 'FP_Pointload', 118.6)
                set_customdata_by_custid(hanger, 7, '12')
            if HostSize == '4"':
                set_parameter_by_name(hanger, 'FP_Pointload', 194.9)
                set_customdata_by_custid(hanger, 7, '20')
            if HostSize == '6"':
                set_parameter_by_name(hanger, 'FP_Pointload', 357.2)
                set_customdata_by_custid(hanger, 7, '36')
            if HostSize == '8"':
                set_parameter_by_name(hanger, 'FP_Pointload', 503.0)
                set_customdata_by_custid(hanger, 7, '50')

        if Hostmat in ['Stainless Steel: 304L', 'Stainless Steel: 316L']:
            HostSize = get_parameter_value_by_name_AsString(doc.GetElement(hosted_info), 'Size')
            if HostSize == '1/2"':
                set_parameter_by_name(hanger, 'FP_Pointload', 4.95)
                set_customdata_by_custid(hanger, 7, '1')
            if HostSize == '3/4"':
                set_parameter_by_name(hanger, 'FP_Pointload', 8.984)
                set_customdata_by_custid(hanger, 7, '2')
            if HostSize == '1"':
                set_parameter_by_name(hanger, 'FP_Pointload', 14.504)
                set_customdata_by_custid(hanger, 7, '2')
            if HostSize == '1 1/4"':
                set_parameter_by_name(hanger, 'FP_Pointload', 25.13)
                set_customdata_by_custid(hanger, 7, '3')
            if HostSize == '1 1/2"':
                set_parameter_by_name(hanger, 'FP_Pointload', 30.47)
                set_customdata_by_custid(hanger, 7, '4')
            if HostSize == '2"':
                set_parameter_by_name(hanger, 'FP_Pointload', 42.2)
                set_customdata_by_custid(hanger, 7, '5')
            if HostSize == '2 1/2"':
                set_parameter_by_name(hanger, 'FP_Pointload', 58.92)
                set_customdata_by_custid(hanger, 7, '6')
            if HostSize == '3"':
                set_parameter_by_name(hanger, 'FP_Pointload', 79.46)
                set_customdata_by_custid(hanger, 7, '8')
            if HostSize == '4"':
                set_parameter_by_name(hanger, 'FP_Pointload', 117.86)
                set_customdata_by_custid(hanger, 7, '12')
            if HostSize == '6"':
                set_parameter_by_name(hanger, 'FP_Pointload', 230.34)
                set_customdata_by_custid(hanger, 7, '24')
            if HostSize == '8"':
                set_parameter_by_name(hanger, 'FP_Pointload', 366.96)
                set_customdata_by_custid(hanger, 7, '37')

    except:
        display_text = "{} (ID: {})".format(family_name, hanger.Id)
        error_data.append((display_text, hanger.Id))

# End Transaction
t.Commit()


class ErrorListForm(Window):
    def __init__(self, error_data, doc, uidoc):
        self.Title = "Hanger Processing Errors"
        self.Width = 400
        self.Height = 450
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.ResizeMode = ResizeMode.CanResize
        self.Topmost = True

        self.doc = doc
        self.uidoc = uidoc

        grid = Grid()
        grid.Margin = Thickness(10)

        grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(2)))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(10)))
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
        for (display_text, eid) in error_data:
            self.listbox.Items.Add(display_text)
        Grid.SetRow(self.listbox, 2)
        grid.Children.Add(self.listbox)

        close_button = Button()
        close_button.Content = "Close"
        close_button.Width = 100
        close_button.Height = 25
        close_button.Margin = Thickness(0, 25, 0, 0)
        close_button.Click += self.close_window
        Grid.SetRow(close_button, 4)
        grid.Children.Add(close_button)

        self.element_map = {display_text: eid for (display_text, eid) in error_data}
        self.listbox.MouseDoubleClick += self.select_element

    def select_element(self, sender, args):
        selected_text = self.listbox.SelectedItem
        if selected_text:
            eid = self.element_map[selected_text]
            element = self.doc.GetElement(eid)
            if element:
                self.uidoc.Selection.SetElementIds(List[ElementId]([eid]))
                self.uidoc.ShowElements(eid)
            else:
                TaskDialog.Show("Error", "Element not found.")

    def close_window(self, sender, args):
        self.Close()


if error_data:
    form = ErrorListForm(error_data, doc, uidoc)
    form.Show()
    while form.IsVisible:
        dispatcher = System.Windows.Threading.Dispatcher.CurrentDispatcher
        dispatcher.Invoke(
            System.Windows.Threading.DispatcherPriority.Background,
            Action(lambda: None)
        )