# -*- coding: utf-8 -*-
from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    FabricationPart, FabricationAncillaryUsage, Transaction, TransactionGroup,
    FilteredElementCollector, BuiltInCategory, ElementId
)
from Autodesk.Revit.UI.Selection import *
from Autodesk.Revit.UI import TaskDialog
from Parameters.Add_SharedParameters import Shared_Params
import os
import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import (
    Application, Window, Thickness, WindowStyle, ResizeMode,
    WindowStartupLocation, HorizontalAlignment, GridLength
)
from System.Windows.Controls import (
    Label, TextBox, Button, Grid, RowDefinition,
    ColumnDefinition, ListBox
)
from System.Windows.Media import Brushes
import System
from System import Action
import System.Windows.Threading
from System.Collections.Generic import List

Shared_Params()

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
    result = integer_part + fractional_part
    return result


def get_parameter_value_by_name(element, parameterName):
    param = element.LookupParameter(parameterName)
    if param and param.HasValue:
        return param.AsString()
    return ""


def get_parameter_value_by_name_AsValueString(element, parameterName):
    param = element.LookupParameter(parameterName)
    if param and param.HasValue:
        return param.AsValueString() or param.AsString()
    return ""


def set_parameter_by_name(element, parameterName, value):
    param = element.LookupParameter(parameterName)
    if param:
        param.Set(value)


def is_trapeze_rack(hanger):
    family_name = get_parameter_value_by_name_AsValueString(hanger, 'Family')
    if not family_name:
        return False

    family_name = family_name.lower()
    trapeze_terms = ['trapeze', 'trapeze rack', 'single strut trapeze', 'double strut trapeze']

    return any(term in family_name for term in trapeze_terms)


class RodSizeForm(Window):
    def __init__(self, initial_values):
        self.Title = "Hanger Rod Sizing"
        self.Width = 300
        self.Height = 335
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.result = None

        grid = Grid()
        grid.Margin = Thickness(10)
        grid.ColumnDefinitions.Add(ColumnDefinition())
        grid.ColumnDefinitions.Add(ColumnDefinition())
        for _ in range(7):
            grid.RowDefinitions.Add(RowDefinition())

        self.Content = grid

        label_texts = [
            '↓ 7/8 Rod / Max Pipe Size:', '↓ 3/4 Rod / Max Pipe Size:',
            '↓ 5/8 Rod / Max Pipe Size:', '↓ 1/2 Rod / Max Pipe Size:',
            '↓ 3/8 Rod / Max Pipe Size:'
        ]
        textbox_names = ['rod_875', 'rod_075', 'rod_625', 'rod_050', 'rod_375']

        self.textboxes = []
        for i, (text, name, value) in enumerate(zip(label_texts, textbox_names, initial_values)):
            label = Label()
            label.Content = text
            label.Margin = Thickness(0, 0, 10, 10)
            Grid.SetRow(label, i)
            Grid.SetColumn(label, 0)
            grid.Children.Add(label)

            textbox = TextBox()
            textbox.Name = name
            textbox.Text = value
            textbox.Width = 100
            textbox.Margin = Thickness(0, 0, 0, 10)
            Grid.SetRow(textbox, i)
            Grid.SetColumn(textbox, 1)
            grid.Children.Add(textbox)
            self.textboxes.append(textbox)

        notice_label = Label()
        notice_label.Content = "Trapeze Rack rod sizes will not be modified"
        notice_label.Margin = Thickness(0, 5, 0, 5)
        notice_label.HorizontalAlignment = HorizontalAlignment.Stretch
        notice_label.HorizontalContentAlignment = HorizontalAlignment.Center
        Grid.SetRow(notice_label, 5)
        Grid.SetColumnSpan(notice_label, 2)
        grid.Children.Add(notice_label)

        ok_button = Button()
        ok_button.Content = "OK"
        ok_button.Width = 60
        ok_button.Height = 25
        ok_button.HorizontalAlignment = HorizontalAlignment.Center
        ok_button.Margin = Thickness(0, 10, 0, 0)
        ok_button.Click += self.on_ok
        Grid.SetRow(ok_button, 6)
        Grid.SetColumnSpan(ok_button, 2)
        grid.Children.Add(ok_button)

        self.values = {}
        if self.textboxes:
            self.textboxes[0].Focus()

    def on_ok(self, sender, args):
        for textbox in self.textboxes:
            self.values[textbox.Name] = textbox.Text
        self.DialogResult = True
        self.Close()


class ErrorListForm(Window):
    def __init__(self, error_data, doc, uidoc):
        self.Title = "Hangers Without Host"
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


try:
    folder_name = "c:\\Temp"
    filepath = os.path.join(folder_name, 'Ribbon_AutoRodSize.txt')

    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
    if not os.path.exists(filepath):
        with open(filepath, 'w') as the_file:
            the_file.writelines(['', '', '', '', ''])

    with open(filepath, 'r') as file:
        lines = [line.rstrip() for line in file.readlines()]

    form = RodSizeForm(lines if lines else ['', '', '', '', ''])
    if form.ShowDialog() and form.DialogResult:
        rod875 = convert_fractions(form.values['rod_875'])
        rod075 = convert_fractions(form.values['rod_075'])
        rod625 = convert_fractions(form.values['rod_625'])
        rod050 = convert_fractions(form.values['rod_050'])
        rod375 = convert_fractions(form.values['rod_375'])

        with open(filepath, 'w') as the_file:
            the_file.writelines([
                str(rod875) + '\n', str(rod075) + '\n',
                str(rod625) + '\n', str(rod050) + '\n',
                str(rod375) + '\n'
            ])

        if rod875 == 0.0:
            rod875 = 0
        if rod075 == 0.0:
            rod075 = 0
        if rod625 == 0.0:
            rod625 = 0
        if rod050 == 0.0:
            rod050 = 0
        if rod375 == 0.0:
            rod375 = 0

        hangers = FilteredElementCollector(doc, curview.Id) \
            .OfCategory(BuiltInCategory.OST_FabricationHangers) \
            .WhereElementIsNotElementType() \
            .ToElements()

        error_data = []

        tg = TransactionGroup(doc, "Change Hanger Rod")
        tg.Start()

        t = Transaction(doc, "Set Hanger Rod")
        t.Start()

        for hanger in hangers:
            if is_trapeze_rack(hanger):
                continue

            try:
                hosted_info = hanger.GetHostedInfo()
                host_id = hosted_info.HostId if hosted_info else ElementId.InvalidElementId

                if not host_id or host_id == ElementId.InvalidElementId:
                    raise Exception("No host")

                host_elem = doc.GetElement(host_id)
                if not host_elem:
                    raise Exception("Host element not found")

                HostSize = convert_fractions(get_parameter_value_by_name(host_elem, 'Size'))

                if HostSize <= rod875:
                    newrodkit = 64
                if HostSize <= rod075:
                    newrodkit = 62
                if HostSize <= rod625:
                    newrodkit = 31
                if HostSize <= rod050:
                    newrodkit = 42
                if HostSize <= rod375:
                    newrodkit = 58

                hanger.HangerRodKit = newrodkit

            except:
                family_name = get_parameter_value_by_name_AsValueString(hanger, 'Family')
                display_text = "{} (ID: {})".format(family_name, hanger.Id)
                error_data.append((display_text, hanger.Id))

        t.Commit()

        t = Transaction(doc, "Update FP Parameter")
        t.Start()
        for x in hangers:
            if is_trapeze_rack(x):
                continue
            for n in x.GetPartAncillaryUsage():
                if n.AncillaryWidthOrDiameter > 0:
                    set_parameter_by_name(x, 'FP_Rod Size', n.AncillaryWidthOrDiameter)
        t.Commit()

        tg.Assimilate()

        if error_data:
            error_form = ErrorListForm(error_data, doc, uidoc)
            error_form.Show()
            while error_form.IsVisible:
                Dispatcher = System.Windows.Threading.Dispatcher.CurrentDispatcher
                Dispatcher.Invoke(
                    System.Windows.Threading.DispatcherPriority.Background,
                    Action(lambda: None)
                )
        else:
            TaskDialog.Show("Success", "Hanger rod sizing completed without missing hosts.")
    else:
        TaskDialog.Show("Cancelled", "Operation cancelled by user.")

except OperationCanceledException:
    TaskDialog.Show("Selection Cancelled", "Operation cancelled by user.")