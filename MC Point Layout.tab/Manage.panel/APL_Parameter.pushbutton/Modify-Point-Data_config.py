# -*- coding: UTF-8 -*-

from Autodesk.Revit.DB import Transaction, FilteredElementCollector, BuiltInCategory
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsDouble
from Parameters.Add_SharedParameters import Shared_Params

import os
import clr
import sys

clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('System')

from System.Windows import Window, Thickness, HorizontalAlignment, ResizeMode, WindowStartupLocation, GridLength
from System.Windows.Controls import Button, TextBox, CheckBox, Grid, RowDefinition, ColumnDefinition, Label

Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView


def feet_to_feet_inches_fraction(decimal_feet, precision=0.125):
    total_inches = decimal_feet * 12
    rounded_total_inches = round(total_inches / precision) * precision
    rounded_total_inches = round(rounded_total_inches, 3)

    feet = int(rounded_total_inches // 12)
    inches = rounded_total_inches % 12
    whole_inches = int(inches)
    fraction = inches - whole_inches

    fraction_str = ""
    if abs(fraction) < 0.01:
        fraction_str = ""
    elif abs(fraction - 0.125) < 0.01:
        fraction_str = "1/8"
    elif abs(fraction - 0.25) < 0.01:
        fraction_str = "1/4"
    elif abs(fraction - 0.375) < 0.01:
        fraction_str = "3/8"
    elif abs(fraction - 0.5) < 0.01:
        fraction_str = "1/2"
    elif abs(fraction - 0.625) < 0.01:
        fraction_str = "5/8"
    elif abs(fraction - 0.75) < 0.01:
        fraction_str = "3/4"
    elif abs(fraction - 0.875) < 0.01:
        fraction_str = "7/8"

    if feet == 0:
        if whole_inches == 0 and fraction_str:
            return "0'-0 {0}\"".format(fraction_str)
        elif whole_inches > 0 and fraction_str:
            return "0'-{0} {1}\"".format(whole_inches, fraction_str)
        elif whole_inches > 0:
            return "0'-{0}\"".format(whole_inches)
        else:
            return "0'-0\""
    else:
        if whole_inches == 0 and fraction_str:
            return "{0}'-0 {1}\"".format(feet, fraction_str)
        elif whole_inches > 0 and fraction_str:
            return "{0}'-{1} {2}\"".format(feet, whole_inches, fraction_str)
        elif whole_inches > 0:
            return "{0}'-{1}\"".format(feet, whole_inches)
        else:
            return "{0}'-0\"".format(feet)


def ensure_settings_file(filepath):
    folder_name = os.path.dirname(filepath)
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)

    if not os.path.exists(filepath):
        with open(filepath, 'w') as f:
            f.writelines(['\n', '\n'])

    try:
        with open(filepath, 'r') as f:
            lines = [line.rstrip() for line in f.readlines()]
    except:
        lines = []

    while len(lines) < 2:
        lines.append("")

    return lines[:2]


def save_settings_file(filepath, desc, pre):
    with open(filepath, 'w') as f:
        f.writelines([desc + '\n', pre + '\n'])


def format_inches_from_feet(value_in_feet):
    return "{0:.2f}".format(value_in_feet * 12)


def get_selected_or_picked_elements():
    selected_ids = list(uidoc.Selection.GetElementIds())
    if selected_ids:
        return [doc.GetElement(el_id) for el_id in selected_ids]

    try:
        picked_refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            'Select Elements or Finish Button'
        )
        return [doc.GetElement(picked_ref.ElementId) for picked_ref in picked_refs]
    except:
        sys.exit()


class UpdateAPLForm(object):
    def __init__(self, desc, pre):
        self._window = Window()
        self._window.Title = "Modify Point Data"
        self._window.Width = 380
        self._window.Height = 190
        self._window.ResizeMode = ResizeMode.NoResize
        self._window.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.values = {}

        grid = Grid()
        grid.Margin = Thickness(12)

        for i in range(4):
            row = RowDefinition()
            row.Height = GridLength.Auto
            grid.RowDefinitions.Add(row)

        col1 = ColumnDefinition()
        col1.Width = GridLength(105)
        grid.ColumnDefinitions.Add(col1)

        col2 = ColumnDefinition()
        col2.Width = GridLength(170)
        grid.ColumnDefinitions.Add(col2)

        col3 = ColumnDefinition()
        col3.Width = GridLength(65)
        grid.ColumnDefinitions.Add(col3)

        row_margin = Thickness(0, 0, 0, 8)
        checkbox_margin = Thickness(0, 2, 0, 8)

        # Prefix
        label_pre = Label()
        label_pre.Content = "Point Prefix:"
        label_pre.Margin = row_margin
        Grid.SetRow(label_pre, 0)
        Grid.SetColumn(label_pre, 0)
        grid.Children.Add(label_pre)

        self.textbox_pre = TextBox()
        self.textbox_pre.Text = pre
        self.textbox_pre.Height = 22
        self.textbox_pre.Width = 135
        self.textbox_pre.IsEnabled = False
        self.textbox_pre.Margin = row_margin
        Grid.SetRow(self.textbox_pre, 0)
        Grid.SetColumn(self.textbox_pre, 1)
        grid.Children.Add(self.textbox_pre)

        enable_pre = Button()
        enable_pre.Content = "Enable"
        enable_pre.Width = 55
        enable_pre.Height = 22
        enable_pre.Margin = row_margin
        enable_pre.HorizontalAlignment = HorizontalAlignment.Center
        enable_pre.Click += lambda sender, args: self.ToggleTextBox(self.textbox_pre, enable_pre)
        Grid.SetRow(enable_pre, 0)
        Grid.SetColumn(enable_pre, 2)
        grid.Children.Add(enable_pre)

        # Description
        label_desc = Label()
        label_desc.Content = "Description:"
        label_desc.Margin = row_margin
        Grid.SetRow(label_desc, 1)
        Grid.SetColumn(label_desc, 0)
        grid.Children.Add(label_desc)

        self.textbox_desc = TextBox()
        self.textbox_desc.Text = desc
        self.textbox_desc.Height = 22
        self.textbox_desc.Width = 135
        self.textbox_desc.IsEnabled = False
        self.textbox_desc.Margin = row_margin
        Grid.SetRow(self.textbox_desc, 1)
        Grid.SetColumn(self.textbox_desc, 1)
        grid.Children.Add(self.textbox_desc)

        enable_desc = Button()
        enable_desc.Content = "Enable"
        enable_desc.Width = 55
        enable_desc.Height = 22
        enable_desc.Margin = row_margin
        enable_desc.HorizontalAlignment = HorizontalAlignment.Center
        enable_desc.Click += lambda sender, args: self.ToggleTextBox(self.textbox_desc, enable_desc)
        Grid.SetRow(enable_desc, 1)
        Grid.SetColumn(enable_desc, 2)
        grid.Children.Add(enable_desc)

        # Sleeve checkbox
        self.checkbox_slv = CheckBox()
        self.checkbox_slv.Content = "Add Size and Length to Sleeve Description (view)"
        self.checkbox_slv.Margin = checkbox_margin
        Grid.SetRow(self.checkbox_slv, 2)
        Grid.SetColumn(self.checkbox_slv, 0)
        Grid.SetColumnSpan(self.checkbox_slv, 3)
        grid.Children.Add(self.checkbox_slv)

        # OK button
        button_ok = Button()
        button_ok.Content = "OK"
        button_ok.Width = 75
        button_ok.Height = 25
        button_ok.Margin = Thickness(0, 6, 0, 0)
        button_ok.HorizontalAlignment = HorizontalAlignment.Center
        Grid.SetRow(button_ok, 3)
        Grid.SetColumn(button_ok, 0)
        Grid.SetColumnSpan(button_ok, 3)
        button_ok.Click += self.OnOK
        grid.Children.Add(button_ok)

        self._window.Content = grid

    def ToggleTextBox(self, textbox, button):
        textbox.IsEnabled = not textbox.IsEnabled
        button.Content = "Disable" if textbox.IsEnabled else "Enable"

    def OnOK(self, sender, args):
        self.values = {
            'Desc': self.textbox_desc.Text,
            'Pre': self.textbox_pre.Text,
            'desc_enabled': self.textbox_desc.IsEnabled,
            'pre_enabled': self.textbox_pre.IsEnabled,
            'writeslvdims': self.checkbox_slv.IsChecked
        }
        self._window.Close()

    def ShowDialog(self):
        self._window.ShowDialog()
        return self.values


try:
    filepath = r"c:\Temp\Ribbon_PointLayout.txt"
    lines = ensure_settings_file(filepath)

    form = UpdateAPLForm(lines[0], lines[1])
    values = form.ShowDialog()

    if not values:
        sys.exit()

    value = values['Desc'].upper()
    value1 = values['Pre'].upper()
    chkdesc = values['desc_enabled']
    chkpre = values['pre_enabled']
    chkslv = values['writeslvdims']

    save_settings_file(filepath, value, value1)

    selection = []
    if chkdesc or chkpre:
        selection = get_selected_or_picked_elements()

    t = Transaction(doc, 'Modify Point Data')
    t.Start()

    try:
        if chkdesc:
            for elem in selection:
                if value:
                    if elem.LookupParameter("TS_Point_Description"):
                        set_parameter_by_name(elem, "TS_Point_Description", value)
                    elif elem.LookupParameter("PointDescription"):
                        set_parameter_by_name(elem, "PointDescription", value)

        if chkpre:
            for elem in selection:
                if value1:
                    if elem.LookupParameter("PointNumber"):
                        set_parameter_by_name(elem, "PointNumber", value1)
                    if elem.LookupParameter("GTP_PointNumber_0"):
                        set_parameter_by_name(elem, "GTP_PointNumber_0", value1)
                    if elem.LookupParameter("GTP_PointNumber_1"):
                        set_parameter_by_name(elem, "GTP_PointNumber_1", value1)
                    if elem.LookupParameter("GTP_PointNumber_2"):
                        set_parameter_by_name(elem, "GTP_PointNumber_2", value1)
                    if elem.LookupParameter("GTP_PointNumber_3"):
                        set_parameter_by_name(elem, "GTP_PointNumber_3", value1)

        if chkslv:
            pipe_accessories = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_PipeAccessory)

            elems = [e for e in pipe_accessories if "Metal Sleeve" in e.Name or "Plastic Sleeve" in e.Name or "Cast Iron Sleeve" in e.Name]
            for x in elems:
                slvdiameter = format_inches_from_feet(get_parameter_value_by_name_AsDouble(x, 'Pipe Nominal Diameter'))
                slvlength = format_inches_from_feet(get_parameter_value_by_name_AsDouble(x, 'Sleeve Length'))
                result_string = 'SLV {0} x {1}'.format(slvdiameter, slvlength)
                set_parameter_by_name(x, 'TS_Point_Description', result_string)

            elems = [e for e in pipe_accessories if "Pipe Riser" in e.Name]
            for x in elems:
                slvdiameter = format_inches_from_feet(get_parameter_value_by_name_AsDouble(x, 'Diameter'))
                result_string = '{0} RISER'.format(slvdiameter)
                set_parameter_by_name(x, 'TS_Point_Description', result_string)

            elems = [e for e in pipe_accessories if "Round Floor Sleeve" in e.Name]
            for x in elems:
                slvdiameter = format_inches_from_feet(get_parameter_value_by_name_AsDouble(x, 'Diameter'))
                slvlength = format_inches_from_feet(get_parameter_value_by_name_AsDouble(x, 'Length'))
                result_string = 'SLV {0} x {1}'.format(slvdiameter, slvlength)
                set_parameter_by_name(x, 'TS_Point_Description', result_string)

            elems = [e for e in pipe_accessories if "Rectangular Sleeve" in e.Name]
            for x in elems:
                slvlength = format_inches_from_feet(get_parameter_value_by_name_AsDouble(x, 'Length'))
                slvwidth = format_inches_from_feet(get_parameter_value_by_name_AsDouble(x, 'Width'))
                slvheight = format_inches_from_feet(get_parameter_value_by_name_AsDouble(x, 'Height'))
                result_string = 'SLV {0} x {1} x {2}'.format(slvlength, slvwidth, slvheight)
                set_parameter_by_name(x, 'TS_Point_Description', result_string)

            elems = [e for e in pipe_accessories if "WS" in e.Name]
            for x in elems:
                slvdiameter = format_inches_from_feet(get_parameter_value_by_name_AsDouble(x, 'Diameter'))
                elevation = get_parameter_value_by_name_AsDouble(x, 'Elevation from Level')
                diameter = get_parameter_value_by_name_AsDouble(x, 'Diameter')
                bottom_elevation = elevation - (diameter / 2)
                slvelevation = feet_to_feet_inches_fraction(bottom_elevation)
                result_string = "DIA {0} BOT {1}".format(slvdiameter, slvelevation)
                set_parameter_by_name(x, 'TS_Point_Description', result_string)

    except Exception as e:
        t.RollBack()
        TaskDialog.Show("Error", "Something went wrong: {0}".format(str(e)))
        sys.exit()

    t.Commit()

except Exception as e:
    TaskDialog.Show("Error", "Failed to execute: {0}".format(str(e)))