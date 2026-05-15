# -*- coding: utf-8 -*-
import os
import sys
import clr
import System

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference("System.Core")

from System import Action
import System.Windows.Threading

from System.Windows import Window, Thickness, HorizontalAlignment, WindowStartupLocation
from System.Windows.Controls import Grid, RowDefinition, ColumnDefinition, Label, TextBox, Button, ListBox, StackPanel
from System.Windows.Interop import WindowInteropHelper
from System.Collections.Generic import List

from Autodesk.Revit.DB import Transaction, FilteredElementCollector, ElementId
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.UI import UIApplication, TaskDialog

from Parameters.Add_SharedParameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name

Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

PARAM_NAME = "FP_REF Line Number"
FOLDER_NAME = r"C:\Temp"
FILE_PATH = os.path.join(FOLDER_NAME, "Ribbon_REFLineNumber.txt")


def natural_key(s):
    import re
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'([0-9]+)', s)]


def ensure_input_file():
    if not os.path.exists(FOLDER_NAME):
        os.makedirs(FOLDER_NAME)

    if not os.path.exists(FILE_PATH):
        with open(FILE_PATH, 'w') as f:
            f.write('123')


def read_previous_input():
    ensure_input_file()
    with open(FILE_PATH, 'r') as f:
        return f.read().strip()


def write_previous_input(value):
    ensure_input_file()
    with open(FILE_PATH, 'w') as f:
        f.write(value)


def get_ref_line_numbers_in_active_view():
    values = set()
    collector = FilteredElementCollector(doc, doc.ActiveView.Id)

    for elem in collector:
        param = elem.LookupParameter(PARAM_NAME)
        if param and param.HasValue:
            val = param.AsString()
            if val:
                values.add(val)

    return sorted(values, key=natural_key)


def get_elements_by_ref_line_number(ref_line_number):
    matches = []
    collector = FilteredElementCollector(doc, doc.ActiveView.Id)

    for elem in collector:
        param = elem.LookupParameter(PARAM_NAME)
        if param and param.HasValue and param.AsString() == ref_line_number:
            matches.append(elem)

    return matches


def show_elements(elements):
    if not elements:
        return

    element_ids = List[ElementId]()
    for elem in elements:
        element_ids.Add(elem.Id)

    uidoc.Selection.SetElementIds(element_ids)
    uidoc.ShowElements(element_ids)


def get_revit_window_handle():
    try:
        return uidoc.Application.MainWindowHandle
    except:
        try:
            return UIApplication(doc.Application).MainWindowHandle
        except:
            return System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle


class REFLineNumberWindow(Window):
    def __init__(self, default_value, ref_line_numbers, revit_window_handle):
        Window.__init__(self)

        self.Title = "REF Line Number"
        self.Width = 300
        self.ResizeMode = System.Windows.ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Topmost = True

        self.accepted = False
        self.result_value = None

        self.InitializeComponents(default_value, ref_line_numbers)

        try:
            WindowInteropHelper(self).Owner = revit_window_handle
        except:
            pass

    def InitializeComponents(self, default_value, ref_line_numbers):
        grid = Grid()
        self.Content = grid

        row_definitions = [
            RowDefinition(Height=System.Windows.GridLength.Auto),
            RowDefinition(Height=System.Windows.GridLength.Auto),
            RowDefinition(Height=System.Windows.GridLength.Auto),
            RowDefinition(Height=System.Windows.GridLength(1, System.Windows.GridUnitType.Star)),
            RowDefinition(Height=System.Windows.GridLength.Auto)
        ]
        for row in row_definitions:
            grid.RowDefinitions.Add(row)

        grid.ColumnDefinitions.Add(
            ColumnDefinition(Width=System.Windows.GridLength(1, System.Windows.GridUnitType.Star))
        )

        item_height = 20
        listbox_height = item_height * min(15, max(7, len(ref_line_numbers))) + 5
        self.Height = listbox_height + 185

        row_index = 0

        self.label = Label()
        self.label.Content = "Enter REF Line Number:"
        self.label.Margin = Thickness(10, 5, 10, 5)
        Grid.SetRow(self.label, row_index)
        grid.Children.Add(self.label)
        row_index += 1

        self.textbox = TextBox()
        self.textbox.Text = default_value
        self.textbox.Margin = Thickness(10, 0, 10, 5)
        Grid.SetRow(self.textbox, row_index)
        grid.Children.Add(self.textbox)
        row_index += 1

        self.list_label = Label()
        self.list_label.Content = "REF Line Numbers in View:"
        self.list_label.Margin = Thickness(10, 0, 10, 5)
        Grid.SetRow(self.list_label, row_index)
        grid.Children.Add(self.list_label)
        row_index += 1

        self.listbox = ListBox()
        self.listbox.Height = listbox_height
        self.listbox.Margin = Thickness(10, 0, 10, 0)

        for number in ref_line_numbers:
            self.listbox.Items.Add(number)

        self.listbox.SelectionChanged += self.on_listbox_select
        self.listbox.MouseDoubleClick += self.on_listbox_double_click

        Grid.SetRow(self.listbox, row_index)
        grid.Children.Add(self.listbox)
        row_index += 1

        button_panel = StackPanel()
        button_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        button_panel.HorizontalAlignment = HorizontalAlignment.Center
        button_panel.Margin = Thickness(0, 15, 0, 10)
        Grid.SetRow(button_panel, row_index)
        grid.Children.Add(button_panel)

        self.ok_button = Button()
        self.ok_button.Content = "OK"
        self.ok_button.Width = 75
        self.ok_button.Height = 25
        self.ok_button.Margin = Thickness(5, 0, 5, 0)
        self.ok_button.Click += self.on_ok_click
        button_panel.Children.Add(self.ok_button)

        self.show_button = Button()
        self.show_button.Content = "Show"
        self.show_button.Width = 75
        self.show_button.Height = 25
        self.show_button.Margin = Thickness(5, 0, 5, 0)
        self.show_button.Click += self.on_show_click
        button_panel.Children.Add(self.show_button)

        self.cancel_button = Button()
        self.cancel_button.Content = "Cancel"
        self.cancel_button.Width = 75
        self.cancel_button.Height = 25
        self.cancel_button.Margin = Thickness(5, 0, 5, 0)
        self.cancel_button.Click += self.on_cancel_click
        button_panel.Children.Add(self.cancel_button)

        self.KeyDown += self.on_key_down
        self.textbox.Focus()
        self.textbox.SelectAll()

    def on_listbox_select(self, sender, event):
        selected = self.listbox.SelectedItem
        if selected:
            self.textbox.Text = str(selected)

    def on_listbox_double_click(self, sender, event):
        selected = self.listbox.SelectedItem
        if not selected:
            return

        self.textbox.Text = str(selected)
        self.show_ref_line_number(str(selected))

    def on_show_click(self, sender, event):
        selected = self.listbox.SelectedItem
        if not selected:
            TaskDialog.Show("Warning", "Please select a REF Line Number from the list.")
            return

        self.textbox.Text = str(selected)
        self.show_ref_line_number(str(selected))

    def show_ref_line_number(self, ref_line_number):
        matching_elements = get_elements_by_ref_line_number(ref_line_number)

        if not matching_elements:
            TaskDialog.Show(
                "Warning",
                "No elements found with REF Line Number '{}' in the active view.".format(ref_line_number)
            )
            return

        show_elements(matching_elements)

    def on_ok_click(self, sender, event):
        self.accepted = True
        self.result_value = self.textbox.Text.strip()
        self.Close()

    def on_cancel_click(self, sender, event):
        self.accepted = False
        self.result_value = None
        self.Close()

    def on_key_down(self, sender, event):
        if event.Key == System.Windows.Input.Key.Enter:
            self.accepted = True
            self.result_value = self.textbox.Text.strip()
            self.Close()
        elif event.Key == System.Windows.Input.Key.Escape:
            self.accepted = False
            self.result_value = None
            self.Close()


previous_input = read_previous_input()
ref_line_numbers = get_ref_line_numbers_in_active_view()
revit_window_handle = get_revit_window_handle()

form = REFLineNumberWindow(previous_input, ref_line_numbers, revit_window_handle)
form.Show()

disp = System.Windows.Threading.Dispatcher.CurrentDispatcher
while form.IsVisible:
    disp.Invoke(System.Windows.Threading.DispatcherPriority.Background, Action(lambda: None))

value = form.result_value if form.accepted else None

if value:
    selected_ids = uidoc.Selection.GetElementIds()

    if not selected_ids:
        try:
            picked_refs = uidoc.Selection.PickObjects(
                ObjectType.Element,
                "Please select elements to set REF Line Number."
            )
            selected_ids = [ref.ElementId for ref in picked_refs]
        except:
            TaskDialog.Show("Error", "Selection cancelled. No elements selected.")
            sys.exit()

    if not selected_ids:
        TaskDialog.Show("Error", "No elements selected. Please select elements and try again.")
        sys.exit()

    selection = [doc.GetElement(eid) for eid in selected_ids]

    write_previous_input(value)

    t = None
    try:
        t = Transaction(doc, 'Set REF Line Number')
        t.Start()

        for elem in selection:
            param_exist = elem.LookupParameter(PARAM_NAME)
            if param_exist and not param_exist.IsReadOnly:
                set_parameter_by_name(elem, PARAM_NAME, value)

        t.Commit()

    except Exception as e:
        if t is not None and t.HasStarted():
            t.RollBack()

        TaskDialog.Show("Error", "Error: {}".format(str(e)))