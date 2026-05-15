import Autodesk
import sys

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    Transaction,
    FabricationConfiguration
)
from Autodesk.Revit.UI import TaskDialog

from Parameters.Add_SharedParameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString

# WPF Imports
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import (
    Window,
    Thickness,
    WindowStyle,
    ResizeMode,
    WindowStartupLocation,
    GridLength
)
from System.Windows.Controls import Label, ListBox, Grid, RowDefinition, Button
from System.Windows.Media import Brushes
import System
from System import Action
import System.Windows.Threading
from System.Collections.Generic import List


Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)
Config = FabricationConfiguration.GetFabricationConfiguration(doc)


def set_customdata_by_custid(fabpart, custid, value):
    fabpart.SetPartCustomDataText(custid, value)


def get_parameter_value_by_name_AsValueString(element, parameterName):
    param = element.LookupParameter(parameterName)
    if param and param.HasValue:
        return param.AsValueString() or param.AsString()
    return ""


def is_trapeze_hanger(hanger):
    family_name = get_parameter_value_by_name_AsValueString(hanger, 'Family') or ""
    return 'trapeze' in family_name.lower()


def set_pointload(hanger, value):
    set_parameter_by_name(hanger, 'FP_Pointload', value)
    set_customdata_by_custid(hanger, 7, str(value))


hanger_collector = (
    FilteredElementCollector(doc, curview.Id)
    .OfCategory(BuiltInCategory.OST_FabricationHangers)
    .WhereElementIsNotElementType()
    .ToElements()
)

error_data = []

t = Transaction(doc, 'Set Pointload Values')
t.Start()

for hanger in hanger_collector:
    family_name = get_parameter_value_by_name_AsValueString(hanger, 'Family') or ""

    # Skip trapeze hangers completely
    if 'trapeze' in family_name.lower():
        continue

    try:
        hosted_info_obj = hanger.GetHostedInfo()

        # Catch non-hosted single hangers and add them to the error list
        if hosted_info_obj is None or hosted_info_obj.HostId == ElementId.InvalidElementId:
            display_text = "{} (ID: {})".format(family_name, hanger.Id)
            error_data.append((display_text, hanger.Id))
            continue

        host_id = hosted_info_obj.HostId
        host = doc.GetElement(host_id)

        if host is None:
            display_text = "{} (ID: {}) - Host Not Found".format(family_name, hanger.Id)
            error_data.append((display_text, hanger.Id))
            continue

        host_mat_param = host.Parameter[BuiltInParameter.FABRICATION_PART_MATERIAL]
        Hostmat = host_mat_param.AsValueString() if host_mat_param else ""
        HostSize = get_parameter_value_by_name_AsString(host, 'Size')

        if Hostmat == 'Cast Iron: Cast Iron':
            if HostSize == '2"':
                set_pointload(hanger, 3)
            elif HostSize == '3"':
                set_pointload(hanger, 5)
            elif HostSize == '4"':
                set_pointload(hanger, 7)
            elif HostSize == '5"':
                set_pointload(hanger, 9)
            elif HostSize == '6"':
                set_pointload(hanger, 12)
            elif HostSize == '8"':
                set_pointload(hanger, 20)
            elif HostSize == '10"':
                set_pointload(hanger, 30)
            elif HostSize == '12"':
                set_pointload(hanger, 42)
            elif HostSize == '15"':
                set_pointload(hanger, 65)

        elif Hostmat == 'Copper: Hard Copper':
            if HostSize == '1/2"':
                set_pointload(hanger, 1)
            elif HostSize == '3/4"':
                set_pointload(hanger, 1)
            elif HostSize == '1"':
                set_pointload(hanger, 2)
            elif HostSize == '1 1/4"':
                set_pointload(hanger, 2)
            elif HostSize == '1 1/2"':
                set_pointload(hanger, 2)
            elif HostSize == '2"':
                set_pointload(hanger, 5)
            elif HostSize == '2 1/2"':
                set_pointload(hanger, 6)
            elif HostSize == '3"':
                set_pointload(hanger, 8)
            elif HostSize == '4"':
                set_pointload(hanger, 15)
            elif HostSize == '6"':
                set_pointload(hanger, 30)
            elif HostSize == '8"':
                set_pointload(hanger, 50)

        elif Hostmat == 'Carbon Steel: Carbon Steel':
            if HostSize == '1/2"':
                set_pointload(hanger, 1)
            elif HostSize == '3/4"':
                set_pointload(hanger, 2)
            elif HostSize == '1"':
                set_pointload(hanger, 3)
            elif HostSize == '1 1/4"':
                set_pointload(hanger, 4)
            elif HostSize == '1 1/2"':
                set_pointload(hanger, 5)
            elif HostSize == '2"':
                set_pointload(hanger, 6)
            elif HostSize == '2 1/2"':
                set_pointload(hanger, 10)
            elif HostSize == '3"':
                set_pointload(hanger, 12)
            elif HostSize == '4"':
                set_pointload(hanger, 20)
            elif HostSize == '6"':
                set_pointload(hanger, 36)
            elif HostSize == '8"':
                set_pointload(hanger, 50)

        elif Hostmat in ['Stainless Steel: 304L', 'Stainless Steel: 316L']:
            if HostSize == '1/2"':
                set_pointload(hanger, 1)
            elif HostSize == '3/4"':
                set_pointload(hanger, 1)
            elif HostSize == '1"':
                set_pointload(hanger, 2)
            elif HostSize == '1 1/4"':
                set_pointload(hanger, 3)
            elif HostSize == '1 1/2"':
                set_pointload(hanger, 4)
            elif HostSize == '2"':
                set_pointload(hanger, 5)
            elif HostSize == '2 1/2"':
                set_pointload(hanger, 6)
            elif HostSize == '3"':
                set_pointload(hanger, 8)
            elif HostSize == '4"':
                set_pointload(hanger, 12)
            elif HostSize == '6"':
                set_pointload(hanger, 24)
            elif HostSize == '8"':
                set_pointload(hanger, 37)

        elif Hostmat in ['PVC: PVC', 'PVC: Sch 40 Clear PVC']:
            if HostSize == '2"':
                set_pointload(hanger, 1)
            elif HostSize == '3"':
                set_pointload(hanger, 2)
            elif HostSize == '4"':
                set_pointload(hanger, 3)
            elif HostSize == '6"':
                set_pointload(hanger, 6)
            elif HostSize == '8"':
                set_pointload(hanger, 9)
            elif HostSize == '10"':
                set_pointload(hanger, 14)

        elif Hostmat.startswith("PolyPro:"):
            if HostSize == '2"':
                set_pointload(hanger, 2)
            elif HostSize == '3"':
                set_pointload(hanger, 3)
            elif HostSize == '4"':
                set_pointload(hanger, 4)
            elif HostSize == '6"':
                set_pointload(hanger, 6)
            elif HostSize == '8"':
                set_pointload(hanger, 10)
            elif HostSize == '10"':
                set_pointload(hanger, 15)

    except Exception as e:
        display_text = "{} (ID: {})".format(family_name, hanger.Id)
        error_data.append((display_text, hanger.Id))

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
        Dispatcher = System.Windows.Threading.Dispatcher.CurrentDispatcher
        Dispatcher.Invoke(
            System.Windows.Threading.DispatcherPriority.Background,
            Action(lambda: None)
        )
else:
    TaskDialog.Show("Success", "Pointload completed on single hangers without error.")