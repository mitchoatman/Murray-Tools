# -*- coding: utf-8 -*-
import os
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

from Autodesk.Revit.DB import FilteredElementCollector, ElementId
from Autodesk.Revit.UI import UIApplication, TaskDialog

from Parameters.Add_SharedParameters import Shared_Params


# -----------------------------------------------------------------------------
# Setup
# -----------------------------------------------------------------------------
Shared_Params()

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
app = UIApplication(doc.Application)

FOLDER_NAME = r"C:\Temp"
FILE_PATH = os.path.join(FOLDER_NAME, "Ribbon_TSPointNumber.txt")
PARAM_NAME = "TS_Point_Number"


# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------
def natural_key(s):
    import re
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'([0-9]+)', s)]


def ensure_input_file():
    if not os.path.exists(FOLDER_NAME):
        os.makedirs(FOLDER_NAME)

    if not os.path.exists(FILE_PATH):
        with open(FILE_PATH, 'w') as f:
            f.write('')


def read_previous_input():
    ensure_input_file()
    with open(FILE_PATH, 'r') as f:
        return f.read().strip()


def write_previous_input(value):
    ensure_input_file()
    with open(FILE_PATH, 'w') as f:
        f.write(value or '')


def get_revit_window_handle():
    try:
        return uidoc.Application.MainWindowHandle
    except:
        try:
            return UIApplication(doc.Application).MainWindowHandle
        except:
            return System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle


def get_point_numbers_in_active_view():
    values = set()
    collector = FilteredElementCollector(doc, doc.ActiveView.Id)

    for elem in collector:
        param = elem.LookupParameter(PARAM_NAME)
        if param and param.HasValue:
            val = param.AsString()
            if val:
                values.add(val)

    return sorted(values, key=natural_key)


def get_elements_by_point_number(point_number):
    matches = []
    collector = FilteredElementCollector(doc, doc.ActiveView.Id)

    for elem in collector:
        param = elem.LookupParameter(PARAM_NAME)
        if param and param.HasValue and param.AsString() == point_number:
            matches.append(elem)

    return matches


def show_elements(elements):
    if not elements:
        return

    ids = List[ElementId]()
    for elem in elements:
        ids.Add(elem.Id)

    uidoc.Selection.SetElementIds(ids)
    uidoc.ShowElements(ids)


# -----------------------------------------------------------------------------
# Window
# -----------------------------------------------------------------------------
class PointNumberWindow(Window):
    def __init__(self, default_value, point_numbers, revit_window_handle):
        Window.__init__(self)

        self.Title = "TS Point Number Locator"
        self.Width = 300
        self.ResizeMode = System.Windows.ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Topmost = True

        self.all_point_numbers = list(point_numbers)

        self.initialize_components(default_value)

        try:
            WindowInteropHelper(self).Owner = revit_window_handle
        except:
            pass

    def initialize_components(self, default_value):
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
        listbox_height = item_height * min(15, max(7, len(self.all_point_numbers))) + 5
        self.Height = listbox_height + 145

        row_index = 0

        self.label = Label()
        self.label.Content = "Search TS Point Number:"
        self.label.Margin = Thickness(10, 5, 10, 5)
        Grid.SetRow(self.label, row_index)
        grid.Children.Add(self.label)
        row_index += 1

        self.textbox = TextBox()
        self.textbox.Text = default_value or ""
        self.textbox.Margin = Thickness(10, 0, 10, 5)
        self.textbox.TextChanged += self.on_text_changed
        Grid.SetRow(self.textbox, row_index)
        grid.Children.Add(self.textbox)
        row_index += 1

        self.list_label = Label()
        self.list_label.Content = "TS Point Numbers in View\nDouble click to zoom:"
        self.list_label.Margin = Thickness(10, 0, 10, 5)
        Grid.SetRow(self.list_label, row_index)
        grid.Children.Add(self.list_label)
        row_index += 1

        self.listbox = ListBox()
        self.listbox.Height = listbox_height
        self.listbox.Margin = Thickness(10, 0, 10, 0)
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

        self.show_button = Button()
        self.show_button.Content = "Show"
        self.show_button.Width = 75
        self.show_button.Height = 25
        self.show_button.Margin = Thickness(5, 0, 5, 0)
        self.show_button.Click += self.on_show_click
        button_panel.Children.Add(self.show_button)

        self.close_button = Button()
        self.close_button.Content = "Close"
        self.close_button.Width = 75
        self.close_button.Height = 25
        self.close_button.Margin = Thickness(5, 0, 5, 0)
        self.close_button.Click += self.on_close_click
        button_panel.Children.Add(self.close_button)

        self.KeyDown += self.on_key_down

        self.refresh_listbox(self.textbox.Text)

    def save_search_text(self):
        try:
            write_previous_input(self.textbox.Text.strip())
        except:
            pass

    def refresh_listbox(self, filter_text=""):
        current_selection = self.listbox.SelectedItem
        self.listbox.Items.Clear()

        filter_text = (filter_text or "").strip().lower()

        filtered = self.all_point_numbers
        if filter_text:
            filtered = [n for n in self.all_point_numbers if filter_text in n.lower()]

        for number in filtered:
            self.listbox.Items.Add(number)

        if current_selection in filtered:
            self.listbox.SelectedItem = current_selection

    def on_text_changed(self, sender, event):
        self.refresh_listbox(self.textbox.Text)

    def on_listbox_double_click(self, sender, event):
        selected = self.listbox.SelectedItem
        if selected:
            self.show_point_number(str(selected))

    def on_show_click(self, sender, event):
        selected = self.listbox.SelectedItem

        if selected:
            self.show_point_number(str(selected))
            return

        typed = self.textbox.Text.strip()
        if typed:
            self.show_point_number(typed)
            return

        TaskDialog.Show("Warning", "Please select a point number from the list or type one in the search box.")

    def show_point_number(self, point_number):
        matching_elements = get_elements_by_point_number(point_number)

        if not matching_elements:
            TaskDialog.Show(
                "Warning",
                "No elements found with point number '{}' in the active view.".format(point_number)
            )
            return

        show_elements(matching_elements)

    def on_close_click(self, sender, event):
        self.save_search_text()
        self.Close()

    def on_key_down(self, sender, event):
        if event.Key == System.Windows.Input.Key.Enter:
            selected = self.listbox.SelectedItem

            if selected:
                self.show_point_number(str(selected))
                return

            typed = self.textbox.Text.strip()
            if typed:
                self.show_point_number(typed)
                return

        elif event.Key == System.Windows.Input.Key.Escape:
            self.save_search_text()
            self.Close()


# -----------------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------------
previous_input = read_previous_input()
point_numbers = get_point_numbers_in_active_view()
revit_window_handle = get_revit_window_handle()

form = PointNumberWindow(previous_input, point_numbers, revit_window_handle)
form.Show()

disp = System.Windows.Threading.Dispatcher.CurrentDispatcher
while form.IsVisible:
    disp.Invoke(System.Windows.Threading.DispatcherPriority.Background, Action(lambda: None))

form.save_search_text()