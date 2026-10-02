# -*- coding: UTF-8 -*-
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, BuiltInCategory
from Autodesk.Revit.UI import TaskDialog, UIThemeManager, UITheme
from Autodesk.Revit.UI.Selection import ObjectType
from Parameters.Get_Set_Params import (
    set_parameter_by_name,
    get_parameter_value_by_name_AsDouble
)
from Parameters.Add_SharedParameters import Shared_Params

import os
import clr
import sys

clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('System')

from System.Windows import (
    Window, Thickness, HorizontalAlignment,
    ResizeMode, WindowStartupLocation, GridLength
)
from System.Windows.Controls import (
    Button, TextBox, CheckBox, Grid,
    RowDefinition, ColumnDefinition, Label
)
from System.Windows.Media import SolidColorBrush, Color as MediaColor

Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView


def set_param_value(element, param_name, value):
    """Set parameter value, checking instance first, then type."""
    param = element.LookupParameter(param_name)
    if param and not param.IsReadOnly:
        param.Set(value)
        return True

    element_type = doc.GetElement(element.GetTypeId())
    if element_type:
        param_type = element_type.LookupParameter(param_name)
        if param_type and not param_type.IsReadOnly:
            param_type.Set(value)
            return True

    return False


def get_family_name(element):
    """Get family name from instance/type, with fallback to element name."""
    try:
        if hasattr(element, "Symbol") and element.Symbol:
            try:
                if element.Symbol.Family and element.Symbol.Family.Name:
                    return element.Symbol.Family.Name.strip()
            except:
                pass

        element_type = doc.GetElement(element.GetTypeId())
        if element_type:
            try:
                if hasattr(element_type, "FamilyName") and element_type.FamilyName:
                    return element_type.FamilyName.strip()
            except:
                pass

        if hasattr(element, "Name") and element.Name:
            return element.Name.strip()
    except:
        pass

    return ""


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
    """Ensure settings file exists and contains at least two lines: desc, pre."""
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
        f.writelines([
            desc + '\n',
            pre + '\n'
        ])


def get_selected_or_picked_elements():
    """Use preselection if available, otherwise prompt user to pick."""
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


def format_inches_from_feet(value_in_feet):
    return "{0:.2f}".format(value_in_feet * 12)


class UpdateAPLForm(object):
    def __init__(self, desc, pre):
        self._window = Window()
        self._window.Title = "Modify Point Data"
        self._window.Width = 360
        self._window.Height = 220
        self._window.ResizeMode = ResizeMode.NoResize
        self._window.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.values = {}

        self.apply_revit_theme()

        grid = Grid()
        grid.Margin = Thickness(12)

        for i in range(5):
            row = RowDefinition()
            row.Height = GridLength.Auto
            grid.RowDefinitions.Add(row)

        col1 = ColumnDefinition()
        col1.Width = GridLength(95)
        grid.ColumnDefinitions.Add(col1)

        col2 = ColumnDefinition()
        col2.Width = GridLength(165)
        grid.ColumnDefinitions.Add(col2)

        col3 = ColumnDefinition()
        col3.Width = GridLength(65)
        grid.ColumnDefinitions.Add(col3)

        row_margin = Thickness(0, 0, 0, 8)
        checkbox_margin = Thickness(0, 2, 0, 8)

        # Prefix
        label_pre = Label()
        label_pre.Content = "Prefix:"
        label_pre.Margin = row_margin
        label_pre.SetResourceReference(Label.ForegroundProperty, "TextForegroundBrush")
        Grid.SetRow(label_pre, 0)
        Grid.SetColumn(label_pre, 0)
        grid.Children.Add(label_pre)

        self.textbox_pre = TextBox()
        self.textbox_pre.Text = pre
        self.textbox_pre.Height = 22
        self.textbox_pre.Width = 135
        self.textbox_pre.IsEnabled = False
        self.textbox_pre.Margin = row_margin
        self.textbox_pre.SetResourceReference(TextBox.BackgroundProperty, "ControlBackgroundBrush")
        self.textbox_pre.SetResourceReference(TextBox.ForegroundProperty, "TextForegroundBrush")
        self.textbox_pre.SetResourceReference(TextBox.BorderBrushProperty, "BorderColorBrush")
        Grid.SetRow(self.textbox_pre, 0)
        Grid.SetColumn(self.textbox_pre, 1)
        grid.Children.Add(self.textbox_pre)

        enable_pre = Button()
        enable_pre.Content = "Enable"
        enable_pre.Width = 55
        enable_pre.Height = 22
        enable_pre.Margin = row_margin
        enable_pre.HorizontalAlignment = HorizontalAlignment.Center
        enable_pre.SetResourceReference(Button.BackgroundProperty, "ButtonBackgroundBrush")
        enable_pre.SetResourceReference(Button.ForegroundProperty, "TextForegroundBrush")
        enable_pre.SetResourceReference(Button.BorderBrushProperty, "BorderColorBrush")
        enable_pre.Click += lambda sender, args: self.ToggleTextBox(self.textbox_pre, enable_pre)
        Grid.SetRow(enable_pre, 0)
        Grid.SetColumn(enable_pre, 2)
        grid.Children.Add(enable_pre)

        # Description
        label_desc = Label()
        label_desc.Content = "Description:"
        label_desc.Margin = row_margin
        label_desc.SetResourceReference(Label.ForegroundProperty, "TextForegroundBrush")
        Grid.SetRow(label_desc, 1)
        Grid.SetColumn(label_desc, 0)
        grid.Children.Add(label_desc)

        self.textbox_desc = TextBox()
        self.textbox_desc.Text = desc
        self.textbox_desc.Height = 22
        self.textbox_desc.Width = 135
        self.textbox_desc.IsEnabled = False
        self.textbox_desc.Margin = row_margin
        self.textbox_desc.SetResourceReference(TextBox.BackgroundProperty, "ControlBackgroundBrush")
        self.textbox_desc.SetResourceReference(TextBox.ForegroundProperty, "TextForegroundBrush")
        self.textbox_desc.SetResourceReference(TextBox.BorderBrushProperty, "BorderColorBrush")
        Grid.SetRow(self.textbox_desc, 1)
        Grid.SetColumn(self.textbox_desc, 1)
        grid.Children.Add(self.textbox_desc)

        enable_desc = Button()
        enable_desc.Content = "Enable"
        enable_desc.Width = 55
        enable_desc.Height = 22
        enable_desc.Margin = row_margin
        enable_desc.HorizontalAlignment = HorizontalAlignment.Center
        enable_desc.SetResourceReference(Button.BackgroundProperty, "ButtonBackgroundBrush")
        enable_desc.SetResourceReference(Button.ForegroundProperty, "TextForegroundBrush")
        enable_desc.SetResourceReference(Button.BorderBrushProperty, "BorderColorBrush")
        enable_desc.Click += lambda sender, args: self.ToggleTextBox(self.textbox_desc, enable_desc)
        Grid.SetRow(enable_desc, 1)
        Grid.SetColumn(enable_desc, 2)
        grid.Children.Add(enable_desc)

        # Sleeve checkbox
        self.checkbox_slv = CheckBox()
        self.checkbox_slv.Content = "Add Size and Length to Sleeve Description"
        self.checkbox_slv.Margin = checkbox_margin
        self.checkbox_slv.SetResourceReference(CheckBox.ForegroundProperty, "TextForegroundBrush")
        Grid.SetRow(self.checkbox_slv, 2)
        Grid.SetColumn(self.checkbox_slv, 0)
        Grid.SetColumnSpan(self.checkbox_slv, 3)
        grid.Children.Add(self.checkbox_slv)

        # Family name checkbox
        self.checkbox_familydesc = CheckBox()
        self.checkbox_familydesc.Content = "Use Family Name for Description"
        self.checkbox_familydesc.Margin = checkbox_margin
        self.checkbox_familydesc.SetResourceReference(CheckBox.ForegroundProperty, "TextForegroundBrush")
        Grid.SetRow(self.checkbox_familydesc, 3)
        Grid.SetColumn(self.checkbox_familydesc, 0)
        Grid.SetColumnSpan(self.checkbox_familydesc, 3)
        grid.Children.Add(self.checkbox_familydesc)

        # OK button
        button_ok = Button()
        button_ok.Content = "OK"
        button_ok.Width = 75
        button_ok.Height = 25
        button_ok.Margin = Thickness(0, 6, 0, 0)
        button_ok.HorizontalAlignment = HorizontalAlignment.Center
        button_ok.SetResourceReference(Button.BackgroundProperty, "ButtonBackgroundBrush")
        button_ok.SetResourceReference(Button.ForegroundProperty, "TextForegroundBrush")
        button_ok.SetResourceReference(Button.BorderBrushProperty, "BorderColorBrush")
        Grid.SetRow(button_ok, 4)
        Grid.SetColumn(button_ok, 0)
        Grid.SetColumnSpan(button_ok, 3)
        button_ok.Click += self.OnOK
        grid.Children.Add(button_ok)

        self._window.Content = grid

    def apply_revit_theme(self):
        try:
            current_theme = UIThemeManager.CurrentTheme
            is_dark = (current_theme == UITheme.Dark)

            bg_hex = "#FF3B4453" if is_dark else "#FFF5F5F5"
            text_hex = "#FFDFDFDF" if is_dark else "#FF333333"
            subtext_hex = "#FF999999" if is_dark else "#FF666666"
            ctrl_bg_hex = "#FF222933" if is_dark else "#FFFFFFFF"
            btn_bg_hex = "#FF222933" if is_dark else "#FFEFEFEF"
            border_hex = "#FF363B40" if is_dark else "#FFD0D0D0"

            res = self._window.Resources
            res["PageBackgroundBrush"] = SolidColorBrush((MediaColor.FromArgb(*[int(bg_hex[i:i+2], 16) for i in (1, 3, 5, 7)])))
            res["TextForegroundBrush"] = SolidColorBrush((MediaColor.FromArgb(*[int(text_hex[i:i+2], 16) for i in (1, 3, 5, 7)])))
            res["SubTextForegroundBrush"] = SolidColorBrush((MediaColor.FromArgb(*[int(subtext_hex[i:i+2], 16) for i in (1, 3, 5, 7)])))
            res["ControlBackgroundBrush"] = SolidColorBrush((MediaColor.FromArgb(*[int(ctrl_bg_hex[i:i+2], 16) for i in (1, 3, 5, 7)])))
            res["ButtonBackgroundBrush"] = SolidColorBrush((MediaColor.FromArgb(*[int(btn_bg_hex[i:i+2], 16) for i in (1, 3, 5, 7)])))
            res["BorderColorBrush"] = SolidColorBrush((MediaColor.FromArgb(*[int(border_hex[i:i+2], 16) for i in (1, 3, 5, 7)])))

            self._window.Background = res["PageBackgroundBrush"]
        except Exception:
            pass

    def ToggleTextBox(self, textbox, button):
        textbox.IsEnabled = not textbox.IsEnabled
        button.Content = "Disable" if textbox.IsEnabled else "Enable"

    def OnOK(self, sender, args):
        self.values = {
            'Desc': self.textbox_desc.Text,
            'Pre': self.textbox_pre.Text,
            'desc_enabled': self.textbox_desc.IsEnabled,
            'pre_enabled': self.textbox_pre.IsEnabled,
            'writeslvdims': self.checkbox_slv.IsChecked,
            'familydesc': self.checkbox_familydesc.IsChecked
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
    desc_enabled = values['desc_enabled']
    pre_enabled = values['pre_enabled']
    chkslv = values['writeslvdims']
    chkfam = values['familydesc']

    save_settings_file(filepath, value, value1)

    selection = []
    if desc_enabled or pre_enabled or chkfam:
        selection = get_selected_or_picked_elements()

    t = Transaction(doc, 'Modify Point Data')
    t.Start()

    try:
        if desc_enabled:
            for elem in selection:
                if value:
                    set_param_value(elem, "PointDescription", value)
                    set_param_value(elem, "TS_Point_Description", value)

        if pre_enabled:
            for elem in selection:
                if value1:
                    set_param_value(elem, "PointNumber", value1)
                    set_param_value(elem, "GTP_PointNumber_0", value1)
                    set_param_value(elem, "GTP_PointNumber_1", value1)
                    set_param_value(elem, "GTP_PointNumber_2", value1)
                    set_param_value(elem, "GTP_PointNumber_3", value1)
                    set_param_value(elem, "TS_Point_Number", value1)

        if chkfam:
            for elem in selection:
                family_name = get_family_name(elem).upper()
                if family_name:
                    set_param_value(elem, "PointDescription", family_name)
                    set_param_value(elem, "TS_Point_Description", family_name)

        if chkslv:
            pipe_accessories = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_PipeAccessory)
            duct_accessories = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_DuctAccessory)

            # Metal / Plastic / Cast Iron Sleeve
            elems = [
                e for e in pipe_accessories
                if "Metal Sleeve" in e.Name or "Plastic Sleeve" in e.Name or "Cast Iron Sleeve" in e.Name
            ]
            for x in elems:
                slvdiameter = format_inches_from_feet(get_parameter_value_by_name_AsDouble(x, 'Pipe Nominal Diameter'))
                slvlength = format_inches_from_feet(get_parameter_value_by_name_AsDouble(x, 'Sleeve Length'))
                result_string = 'SLV {0} x {1}'.format(slvdiameter, slvlength)
                set_parameter_by_name(x, 'TS_Point_Description', result_string)

            # Pipe Riser
            elems = [e for e in pipe_accessories if "Pipe Riser" in e.Name]
            for x in elems:
                slvdiameter = format_inches_from_feet(get_parameter_value_by_name_AsDouble(x, 'Diameter'))
                result_string = "{0} RISER".format(slvdiameter)
                set_parameter_by_name(x, 'TS_Point_Description', result_string)

            # Floor Sleeve
            elems = [e for e in pipe_accessories if "Floor Sleeve" in e.Name or "Round Floor Sleeve" in e.Name]
            for x in elems:
                slvdiameter = format_inches_from_feet(get_parameter_value_by_name_AsDouble(x, 'Diameter'))
                slvlength = format_inches_from_feet(get_parameter_value_by_name_AsDouble(x, 'Length'))
                result_string = 'SLV {0} x {1}'.format(slvdiameter, slvlength)
                set_parameter_by_name(x, 'TS_Point_Description', result_string)

            # Rectangular Sleeve - Pipe Accessory
            elems = [e for e in pipe_accessories if "Rectangular Sleeve" in e.Name]
            for x in elems:
                slvlength = format_inches_from_feet(get_parameter_value_by_name_AsDouble(x, 'Length'))
                slvwidth = format_inches_from_feet(get_parameter_value_by_name_AsDouble(x, 'Width'))
                slvheight = format_inches_from_feet(get_parameter_value_by_name_AsDouble(x, 'Height'))
                result_string = 'SLV {0} x {1} x {2}'.format(slvlength, slvwidth, slvheight)
                set_parameter_by_name(x, 'TS_Point_Description', result_string)

            # WS / DR-WS
            elems = [e for e in pipe_accessories if "WS" in e.Name]
            for x in elems:
                slvdiameter = format_inches_from_feet(get_parameter_value_by_name_AsDouble(x, 'Diameter'))
                slvelevation = feet_to_feet_inches_fraction(
                    get_parameter_value_by_name_AsDouble(x, 'Elevation from Level')
                )
                result_string = "DIA {0} CL {1}".format(slvdiameter, slvelevation)
                set_parameter_by_name(x, 'TS_Point_Description', result_string)

            # Blockout
            elems = [e for e in pipe_accessories if "BLOCKOUT" in e.Name]
            for x in elems:
                slvlength = format_inches_from_feet(get_parameter_value_by_name_AsDouble(x, 'Length'))
                slvwidth = format_inches_from_feet(get_parameter_value_by_name_AsDouble(x, 'Width'))
                slvheight = format_inches_from_feet(get_parameter_value_by_name_AsDouble(x, 'Height'))
                result_string = 'BO W{0} x H{1} x L{2}'.format(slvwidth, slvheight, slvlength)
                set_parameter_by_name(x, 'TS_Point_Description', result_string)

            # RWS - Duct Accessory
            elems = [e for e in duct_accessories if "RWS" in e.Name]
            for x in elems:
                slvlength = format_inches_from_feet(get_parameter_value_by_name_AsDouble(x, 'Length'))
                slvwidth = format_inches_from_feet(get_parameter_value_by_name_AsDouble(x, 'Width'))
                slvheight = format_inches_from_feet(get_parameter_value_by_name_AsDouble(x, 'Height'))
                result_string = 'RWS W{0} x H{1} x L{2}'.format(slvwidth, slvheight, slvlength)
                set_parameter_by_name(x, 'TS_Point_Description', result_string)

            # Rectangular Sleeve - Duct Accessory
            elems = [e for e in duct_accessories if "Rectangular Sleeve" in e.Name]
            for x in elems:
                slvlength = format_inches_from_feet(get_parameter_value_by_name_AsDouble(x, 'Length'))
                slvwidth = format_inches_from_feet(get_parameter_value_by_name_AsDouble(x, 'Width'))
                slvheight = format_inches_from_feet(get_parameter_value_by_name_AsDouble(x, 'Height'))
                result_string = 'SLV {0} x {1} x {2}'.format(slvlength, slvwidth, slvheight)
                set_parameter_by_name(x, 'TS_Point_Description', result_string)

    except Exception as e:
        t.RollBack()
        TaskDialog.Show("Error", "Something went wrong: {0}".format(str(e)))
        sys.exit()

    t.Commit()

except Exception as e:
    TaskDialog.Show("Error", "Failed to execute: {0}".format(str(e)))