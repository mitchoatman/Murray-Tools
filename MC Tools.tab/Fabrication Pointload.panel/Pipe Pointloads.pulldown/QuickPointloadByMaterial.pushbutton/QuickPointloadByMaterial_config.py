# -*- coding: UTF-8 -*-
import Autodesk
import sys
import clr

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    Transaction,
    FabricationConfiguration
)
from pyrevit import revit, DB, script, forms
from Parameters.Add_SharedParameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString

# WPF & Windows Forms Imports
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from System.Windows import Window, Thickness, WindowStyle, ResizeMode, WindowStartupLocation, GridLength
from System.Windows.Controls import Label, ListBox, Grid, RowDefinition, Button
from System.Windows.Media import Brushes
import System
from System import Action
import System.Windows.Threading
from System.Collections.Generic import List
from Autodesk.Revit.UI import TaskDialog
from System.Windows.Forms import NotifyIcon, ToolTipIcon
from System.Drawing import SystemIcons

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


def show_balloon_notification(title, message, timeout=5000):
    """Displays a native Windows balloon notification in the system tray area."""
    notify_icon = NotifyIcon()
    try:
        notify_icon.Icon = SystemIcons.Information
        notify_icon.Visible = True
        notify_icon.ShowBalloonTip(timeout, title, message, ToolTipIcon.Info)
    except Exception:
        pass


# This writes to fab part custom data field
def set_customdata_by_custid(fabpart, custid, value):
    fabpart.SetPartCustomDataText(custid, value)


def get_parameter_value_by_name_AsValueString(element, parameterName):
    param = element.LookupParameter(parameterName)
    if param and param.HasValue:
        return param.AsValueString() or param.AsString()
    return ""


def set_pointload(hanger, load_value, rounded_value):
    set_parameter_by_name(hanger, 'FP_Pointload', load_value)
    set_customdata_by_custid(hanger, 7, str(rounded_value))


hanger_collector = (
    FilteredElementCollector(doc, curview.Id)
    .OfCategory(BuiltInCategory.OST_FabricationHangers)
    .WhereElementIsNotElementType()
    .ToElements()
)

error_data = []

t = Transaction(doc, 'Set Full Pointload Values')
t.Start()

for hanger in hanger_collector:
    family_name = get_parameter_value_by_name_AsValueString(hanger, 'Family') or ""

    # Skip trapeze hangers
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
                set_pointload(hanger, 24.75, 3)
            elif HostSize == '3"':
                set_pointload(hanger, 41.2, 5)
            elif HostSize == '4"':
                set_pointload(hanger, 64.1, 7)
            elif HostSize == '5"':
                set_pointload(hanger, 87.5, 9)
            elif HostSize == '6"':
                set_pointload(hanger, 115.9, 12)
            elif HostSize == '8"':
                set_pointload(hanger, 198.3, 20)
            elif HostSize == '10"':
                set_pointload(hanger, 298.0, 30)
            elif HostSize == '12"':
                set_pointload(hanger, 420.0, 42)
            elif HostSize == '15"':
                set_pointload(hanger, 650.0, 65)

        elif Hostmat == 'Copper: Hard Copper':
            if HostSize == '1/2"':
                set_pointload(hanger, 2.638, 1)
            elif HostSize == '3/4"':
                set_pointload(hanger, 7.56, 1)
            elif HostSize == '1"':
                set_pointload(hanger, 10.68, 2)
            elif HostSize == '1 1/4"':
                set_pointload(hanger, 11.58, 2)
            elif HostSize == '1 1/2"':
                set_pointload(hanger, 16.56, 2)
            elif HostSize == '2"':
                set_pointload(hanger, 42.1, 5)
            elif HostSize == '2 1/2"':
                set_pointload(hanger, 58.8, 6)
            elif HostSize == '3"':
                set_pointload(hanger, 80.3, 8)
            elif HostSize == '4"':
                set_pointload(hanger, 147.5, 15)
            elif HostSize == '6"':
                set_pointload(hanger, 292.8, 30)
            elif HostSize == '8"':
                set_pointload(hanger, 500.0, 50)

        elif Hostmat == 'Carbon Steel: Carbon Steel':
            if HostSize == '1/2"':
                set_pointload(hanger, 8.7, 1)
            elif HostSize == '3/4"':
                set_pointload(hanger, 14.32, 2)
            elif HostSize == '1"':
                set_pointload(hanger, 21.36, 3)
            elif HostSize == '1 1/4"':
                set_pointload(hanger, 36.2, 4)
            elif HostSize == '1 1/2"':
                set_pointload(hanger, 42.4, 5)
            elif HostSize == '2"':
                set_pointload(hanger, 58.9, 6)
            elif HostSize == '2 1/2"':
                set_pointload(hanger, 91.0, 10)
            elif HostSize == '3"':
                set_pointload(hanger, 118.6, 12)
            elif HostSize == '4"':
                set_pointload(hanger, 194.9, 20)
            elif HostSize == '6"':
                set_pointload(hanger, 357.2, 36)
            elif HostSize == '8"':
                set_pointload(hanger, 503.0, 50)

        elif Hostmat in ['Stainless Steel: 304L', 'Stainless Steel: 316L']:
            if HostSize == '1/2"':
                set_pointload(hanger, 4.95, 1)
            elif HostSize == '3/4"':
                set_pointload(hanger, 8.984, 2)
            elif HostSize == '1"':
                set_pointload(hanger, 14.504, 2)
            elif HostSize == '1 1/4"':
                set_pointload(hanger, 25.13, 3)
            elif HostSize == '1 1/2"':
                set_pointload(hanger, 30.47, 4)
            elif HostSize == '2"':
                set_pointload(hanger, 42.2, 5)
            elif HostSize == '2 1/2"':
                set_pointload(hanger, 58.92, 6)
            elif HostSize == '3"':
                set_pointload(hanger, 79.46, 8)
            elif HostSize == '4"':
                set_pointload(hanger, 117.86, 12)
            elif HostSize == '6"':
                set_pointload(hanger, 230.34, 24)
            elif HostSize == '8"':
                set_pointload(hanger, 366.96, 37)

        elif Hostmat in ['PVC: PVC', 'PVC: Sch 40 Clear PVC', 'PVC: CPVC']:
            if HostSize == '2"':
                set_pointload(hanger, 8.4, 1)
            elif HostSize == '3"':
                set_pointload(hanger, 14.0, 2)
            elif HostSize == '4"':
                set_pointload(hanger, 24.0, 3)
            elif HostSize == '6"':
                set_pointload(hanger, 51.0, 6)
            elif HostSize == '8"':
                set_pointload(hanger, 89.0, 9)
            elif HostSize == '10"':
                set_pointload(hanger, 138.0, 14)

        # Optional PolyPro block if you want to add full weights later
        # elif Hostmat.startswith('PolyPro:'):
        #     if HostSize == '2"':
        #         set_pointload(hanger, YOUR_VALUE_HERE, 2)
        #     elif HostSize == '3"':
        #         set_pointload(hanger, YOUR_VALUE_HERE, 3)

    except Exception:
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
        dispatcher = System.Windows.Threading.Dispatcher.CurrentDispatcher
        dispatcher.Invoke(
            System.Windows.Threading.DispatcherPriority.Background,
            Action(lambda: None)
        )
else:
    show_balloon_notification("Success", "Pointload completed on single hangers without error.")