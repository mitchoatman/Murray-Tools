# -*- coding: utf-8 -*-
import clr
clr.AddReference('System')
import System
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import Window, Thickness, HorizontalAlignment, WindowStartupLocation, ResizeMode
from System.Windows.Controls import Grid, RowDefinition, Label, TextBox, Button, StackPanel, Orientation
from System.Windows.Interop import WindowInteropHelper

from Autodesk.Revit import DB
from Autodesk.Revit.DB import XYZ, Transaction
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

import os

# Current Revit document and UI document
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
uiapp = __revit__

# File path for saving distances
folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_MoveGeneric.txt')


class GenericModelFilter(ISelectionFilter):
    def AllowElement(self, element):
        return element.Category and element.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_GenericModel)

    def AllowReference(self, reference, point):
        return False


class MoveGenericWindow(Window):
    def __init__(self, revit_window_handle, default_x, default_y):
        Window.__init__(self)

        self.Title = "Move Generic Model Elements"
        self.Width = 420
        self.Height = 220
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen

        self.x_dist = default_x
        self.y_dist = default_y
        self.confirmed = False

        self.InitializeComponents()
        WindowInteropHelper(self).Owner = revit_window_handle

    def InitializeComponents(self):
        grid = Grid()
        self.Content = grid

        row_definitions = [
            RowDefinition(Height=System.Windows.GridLength.Auto),
            RowDefinition(Height=System.Windows.GridLength.Auto),
            RowDefinition(Height=System.Windows.GridLength.Auto),
            RowDefinition(Height=System.Windows.GridLength.Auto),
            RowDefinition(Height=System.Windows.GridLength.Auto)
        ]

        for row in row_definitions:
            grid.RowDefinitions.Add(row)

        row_index = 0

        label_x = Label()
        label_x.Content = "X Distance (ft) (positive = right, negative = left):"
        label_x.Margin = Thickness(10, 10, 10, 2)
        Grid.SetRow(label_x, row_index)
        grid.Children.Add(label_x)
        row_index += 1

        self.textbox_x = TextBox()
        self.textbox_x.Text = str(self.x_dist)
        self.textbox_x.Margin = Thickness(10, 0, 10, 8)
        Grid.SetRow(self.textbox_x, row_index)
        grid.Children.Add(self.textbox_x)
        row_index += 1

        label_y = Label()
        label_y.Content = "Y Distance (ft) (positive = up, negative = down):"
        label_y.Margin = Thickness(10, 5, 10, 2)
        Grid.SetRow(label_y, row_index)
        grid.Children.Add(label_y)
        row_index += 1

        self.textbox_y = TextBox()
        self.textbox_y.Text = str(self.y_dist)
        self.textbox_y.Margin = Thickness(10, 0, 10, 8)
        Grid.SetRow(self.textbox_y, row_index)
        grid.Children.Add(self.textbox_y)
        row_index += 1

        button_panel = StackPanel()
        button_panel.Orientation = Orientation.Horizontal
        button_panel.HorizontalAlignment = HorizontalAlignment.Center
        button_panel.Margin = Thickness(0, 10, 0, 10)
        Grid.SetRow(button_panel, row_index)
        grid.Children.Add(button_panel)

        self.ok_button = Button()
        self.ok_button.Content = "OK"
        self.ok_button.Width = 75
        self.ok_button.Height = 25
        self.ok_button.Margin = Thickness(5, 0, 5, 0)
        self.ok_button.Click += self.on_ok_click
        button_panel.Children.Add(self.ok_button)

        self.cancel_button = Button()
        self.cancel_button.Content = "Cancel"
        self.cancel_button.Width = 75
        self.cancel_button.Height = 25
        self.cancel_button.Margin = Thickness(5, 0, 5, 0)
        self.cancel_button.Click += self.on_cancel_click
        button_panel.Children.Add(self.cancel_button)

    def on_ok_click(self, sender, event):
        x_text = self.textbox_x.Text.strip()
        y_text = self.textbox_y.Text.strip()

        try:
            float(x_text)
            float(y_text)
        except:
            TaskDialog.Show("Error", "Please enter valid numeric values for distances.")
            return

        self.x_dist = x_text
        self.y_dist = y_text
        self.confirmed = True
        self.Close()

    def on_cancel_click(self, sender, event):
        self.confirmed = False
        self.Close()


def pick_elements():
    try:
        selection_filter = GenericModelFilter()
        selections = uidoc.Selection.PickObjects(
            ObjectType.Element,
            selection_filter,
            "Window select generic model elements to move (drag from left to right)."
        )

        if selections:
            return [doc.GetElement(sel.ElementId) for sel in selections]
        else:
            TaskDialog.Show("Move Generic Model Elements", "No generic model elements selected. Script will exit.")
            return None

    except:
        TaskDialog.Show("Move Generic Model Elements", "Selection cancelled or failed. Script will exit.")
        return None


def load_saved_distances():
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)

    if not os.path.exists(filepath):
        with open(filepath, 'w') as the_file:
            the_file.writelines(["0\n", "0\n"])

    with open(filepath, 'r') as file:
        lines = [line.rstrip() for line in file.readlines()]

    if len(lines) < 2:
        with open(filepath, 'w') as the_file:
            the_file.writelines(["0\n", "0\n"])
        return "0", "0"

    return lines[0], lines[1]


def save_distances(x_dist, y_dist):
    with open(filepath, 'w') as file:
        file.write("{}\n{}\n".format(x_dist, y_dist))


def move_elements():
    elements = pick_elements()
    if not elements:
        return

    saved_x_dist, saved_y_dist = load_saved_distances()

    revit_window_handle = uiapp.MainWindowHandle
    form = MoveGenericWindow(revit_window_handle, saved_x_dist, saved_y_dist)
    form.ShowDialog()

    if not form.confirmed:
        return

    try:
        x_dist = float(form.x_dist)
        y_dist = float(form.y_dist)
    except:
        TaskDialog.Show("Error", "Please enter valid numeric values for distances.")
        return

    save_distances(x_dist, y_dist)

    t = Transaction(doc, "Move Generic Model Elements")
    t.Start()
    try:
        translation = XYZ(x_dist, y_dist, 0)
        for element in elements:
            DB.ElementTransformUtils.MoveElement(doc, element.Id, translation)
        t.Commit()
    except Exception as e:
        if t.HasStarted():
            t.RollBack()
        TaskDialog.Show("Error", "Error moving elements: {}".format(str(e)))


if __name__ == "__main__":
    move_elements()