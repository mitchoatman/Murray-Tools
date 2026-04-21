# -*- coding: utf-8 -*-
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FabricationPart, FabricationAncillaryUsage, Transaction, TransactionGroup, FilteredElementCollector, BuiltInCategory
from Autodesk.Revit.UI.Selection import *
from Autodesk.Revit.UI import TaskDialog
from Parameters.Add_SharedParameters import Shared_Params
import os
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
from System.Windows import Application, Window, Thickness, WindowStyle, ResizeMode, WindowStartupLocation, HorizontalAlignment
from System.Windows.Controls import Label, TextBox, Button, Grid, RowDefinition, ColumnDefinition

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
    return element.LookupParameter(parameterName).AsString()

def get_parameter_value_by_name_AsValueString(element, parameterName):
    return element.LookupParameter(parameterName).AsValueString()

def set_parameter_by_name(element, parameterName, value):
    element.LookupParameter(parameterName).Set(value)

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

    class RodSizeForm(Window):
        def __init__(self, initial_values):
            self.Title = "Hanger Rod Sizing"
            self.Width = 300
            self.Height = 300
            self.WindowStyle = WindowStyle.SingleBorderWindow
            self.ResizeMode = ResizeMode.NoResize
            self.WindowStartupLocation = WindowStartupLocation.CenterScreen
            self.result = None

            grid = Grid()
            grid.Margin = Thickness(10)
            grid.ColumnDefinitions.Add(ColumnDefinition())
            grid.ColumnDefinitions.Add(ColumnDefinition())
            for _ in range(6):
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

            ok_button = Button()
            ok_button.Content = "OK"
            ok_button.Width = 60
            ok_button.Height = 25
            ok_button.HorizontalAlignment = HorizontalAlignment.Center
            ok_button.Margin = Thickness(0, 10, 0, 0)
            ok_button.Click += self.on_ok
            Grid.SetRow(ok_button, 5)
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

        if rod875 == 0.0 or '':
            rod875 = 0
        if rod075 == 0.0 or '':
            rod075 = 0
        if rod625 == 0.0 or '':
            rod625 = 0
        if rod050 == 0.0 or '':
            rod050 = 0
        if rod375 == 0.0 or '':
            rod375 = 0

        hangers = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationHangers) \
                           .WhereElementIsNotElementType() \
                           .ToElements()

        tg = TransactionGroup(doc, "Change Hanger Rod")
        tg.Start()

        t = Transaction(doc, "Set Hanger Rod")
        t.Start()
        hangers_without_host_printed = False
        for hanger in hangers:
            hosted_info = hanger.GetHostedInfo().HostId
            try:
                HostSize = convert_fractions(get_parameter_value_by_name(doc.GetElement(hosted_info), 'Size'))
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
                from pyrevit import script
                output = script.get_output()
                print('{}: {}'.format((get_parameter_value_by_name_AsValueString(hanger, 'Family')), output.linkify(hanger.Id)))
        t.Commit()

        t = Transaction(doc, "Update FP Parameter")
        t.Start()
        for x in hangers:
            [set_parameter_by_name(x, 'FP_Rod Size', n.AncillaryWidthOrDiameter) for n in x.GetPartAncillaryUsage() if n.AncillaryWidthOrDiameter > 0]
        t.Commit()

        tg.Assimilate()
    else:
        TaskDialog.Show("Cancelled", "Operation cancelled by user.")
except OperationCanceledException:
    TaskDialog.Show("Selection Cancelled", "Operation cancelled by user.")