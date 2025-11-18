# -*- coding: UTF-8 -*-
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, BuiltInCategory, ElementId
from Autodesk.Revit.UI import Selection, TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString, get_parameter_value_by_name_AsDouble
from Parameters.Add_SharedParameters import Shared_Params
import os, clr, sys
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('System')
from System.Windows import Application, Window, Thickness, HorizontalAlignment, ResizeMode, WindowStartupLocation
from System.Windows.Controls import Button, TextBox, CheckBox, Grid, RowDefinition, ColumnDefinition, Label

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

    # If instance param not found, try type
    element_type = doc.GetElement(element.GetTypeId())
    if element_type:
        param_type = element_type.LookupParameter(param_name)
        if param_type and not param_type.IsReadOnly:
            param_type.Set(value)
            return True

    return False


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

try:
    folder_name = "c:\\Temp"
    filepath = os.path.join(folder_name, 'Ribbon_PointLayout.txt')

    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
    if not os.path.exists(filepath):
        with open(filepath, 'w') as f:
            f.write('123')

    try:
        with open(filepath, 'r') as file:
            lines = file.readlines()
            lines = [line.rstrip() for line in lines]
    except:
        with open(filepath, 'w') as the_file:
            line1 = 'pre\n'
            line2 = 'num\n'
            the_file.writelines([line1, line2]) 

    if len(lines) < 2:
        with open(filepath, 'w') as the_file:
            line1 = 'desc\n'
            line2 = 'pre\n'
            the_file.writelines([line1, line2]) 

    with open(filepath, 'r') as file:
        lines = file.readlines()
        lines = [line.rstrip() for line in lines]

    class UpdateAPLForm(object):
        def __init__(self, desc, pre):
            self._window = Window()
            self._window.Title = "Update APL Information"
            self._window.Width = 360
            self._window.Height = 230
            self._window.ResizeMode = ResizeMode.NoResize
            self._window.WindowStartupLocation = WindowStartupLocation.CenterScreen
            self.values = {}

            # Create grid layout
            grid = Grid()
            grid.Margin = Thickness(10)
            for i in range(5):  # 5 rows: 2 for textboxes/buttons, 2 for checkboxes, 1 for OK button
                grid.RowDefinitions.Add(RowDefinition())
            grid.ColumnDefinitions.Add(ColumnDefinition())  # Label column
            grid.ColumnDefinitions.Add(ColumnDefinition())  # TextBox column
            grid.ColumnDefinitions.Add(ColumnDefinition())  # Enable button column

            # Prefix label and textbox
            label_pre = Label()
            label_pre.Content = "Prefix:"
            Grid.SetRow(label_pre, 0)
            Grid.SetColumn(label_pre, 0)
            grid.Children.Add(label_pre)

            self.textbox_pre = TextBox()
            self.textbox_pre.Text = pre
            self.textbox_pre.Height = 20
            self.textbox_pre.Width = 135
            self.textbox_pre.IsEnabled = False
            Grid.SetRow(self.textbox_pre, 0)
            Grid.SetColumn(self.textbox_pre, 1)
            grid.Children.Add(self.textbox_pre)

            # Prefix enable button
            enable_pre = Button()
            enable_pre.Content = "Enable"
            enable_pre.Width = 50
            enable_pre.Height = 20
            enable_pre.Margin = Thickness(-15, 0, 20, 0)
            enable_pre.HorizontalAlignment = HorizontalAlignment.Right
            enable_pre.Click += lambda sender, args: self.ToggleTextBox(self.textbox_pre, enable_pre)
            Grid.SetRow(enable_pre, 0)
            Grid.SetColumn(enable_pre, 2)
            grid.Children.Add(enable_pre)

            # Description label and textbox
            label_desc = Label()
            label_desc.Content = "Description:"
            Grid.SetRow(label_desc, 1)
            Grid.SetColumn(label_desc, 0)
            grid.Children.Add(label_desc)

            self.textbox_desc = TextBox()
            self.textbox_desc.Text = desc
            self.textbox_desc.Height = 20
            self.textbox_desc.Width = 135
            self.textbox_desc.IsEnabled = False
            Grid.SetRow(self.textbox_desc, 1)
            Grid.SetColumn(self.textbox_desc, 1)
            grid.Children.Add(self.textbox_desc)

            # Description enable button
            enable_desc = Button()
            enable_desc.Content = "Enable"
            enable_desc.Width = 50
            enable_desc.Height = 20
            enable_desc.Margin = Thickness(-15, 0, 20, 0)
            enable_desc.HorizontalAlignment = HorizontalAlignment.Right
            enable_desc.Click += lambda sender, args: self.ToggleTextBox(self.textbox_desc, enable_desc)
            Grid.SetRow(enable_desc, 1)
            Grid.SetColumn(enable_desc, 2)
            grid.Children.Add(enable_desc)

            # Checkboxes
            self.checkbox_cleanup = CheckBox()
            self.checkbox_cleanup.Content = "Re-format Insert Description"
            Grid.SetRow(self.checkbox_cleanup, 2)
            Grid.SetColumn(self.checkbox_cleanup, 0)
            Grid.SetColumnSpan(self.checkbox_cleanup, 3)
            grid.Children.Add(self.checkbox_cleanup)

            self.checkbox_slv = CheckBox()
            self.checkbox_slv.Content = "Add Size and Length to Sleeve Description"
            self.checkbox_slv.Margin = Thickness(0, -5, 0, 0)
            Grid.SetRow(self.checkbox_slv, 3)
            Grid.SetColumn(self.checkbox_slv, 0)
            Grid.SetColumnSpan(self.checkbox_slv, 3)
            grid.Children.Add(self.checkbox_slv)

            # OK Button
            button_ok = Button()
            button_ok.Content = "OK"
            button_ok.Width = 75
            button_ok.Height = 25
            button_ok.Margin = Thickness(0, -5, 0, 0)
            button_ok.HorizontalAlignment = HorizontalAlignment.Center
            Grid.SetRow(button_ok, 4)
            Grid.SetColumn(button_ok, 0)
            Grid.SetColumnSpan(button_ok, 3)  # Span all columns for centering
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
                'cleanupins': self.checkbox_cleanup.IsChecked,
                'writeslvdims': self.checkbox_slv.IsChecked
            }
            self._window.Close()

        def ShowDialog(self):
            self._window.ShowDialog()
            return self.values

    # Display dialog
    form = UpdateAPLForm(lines[0], lines[1])
    values = form.ShowDialog()
    if not values:
        sys.exit()  # Exit quietly if dialog is closed with X or cancelled

    # Convert dialog input into variables
    value = values['Desc'].upper()
    value1 = values['Pre'].upper()
    desc_enabled = values['desc_enabled']
    pre_enabled = values['pre_enabled']
    chkins = values['cleanupins']
    chkslv = values['writeslvdims']

    # Write values to text file
    with open(filepath, 'w') as the_file:
        line1 = value + '\n'
        line2 = value1 + '\n'
        the_file.writelines([line1, line2])

    # Prompt for selection only if description or prefix textboxes are enabled
    selection = []
    if desc_enabled or pre_enabled:
        OBJselection = uidoc.Selection.PickObjects(ObjectType.Element, 'Select Elements or Finish Button')
        selection = [doc.GetElement(elId) for elId in OBJselection]

    t = Transaction(doc, 'Modify Point Data')
    t.Start()

    try:
        if desc_enabled:
            for i in selection:
                if value:
                    if not set_param_value(i, "TS_Point_Description", value):
                        set_param_value(i, "PointDescription", value)
        if pre_enabled:
            for i in selection:
                if value1:
                    set_param_value(i, "PointNumber", value1)
                    set_param_value(i, "GTP_PointNumber_0", value1)
                    set_param_value(i, "GTP_PointNumber_1", value1)
                    set_param_value(i, "GTP_PointNumber_2", value1)
                    set_param_value(i, "GTP_PointNumber_3", value1)
        if chkins:
            generic_models_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_GenericModel)
            pipe_point_elements = [element for element in generic_models_collector if "Pipe Pt" in element.Name]
            for x in pipe_point_elements:
                originalvalue = get_parameter_value_by_name_AsString(x, 'PointDescription')
                result_string = originalvalue.replace("0' - 0 ", "")
                set_parameter_by_name(x, 'PointDescription', result_string)
        if chkslv:
            accessory_models_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_PipeAccessory)
            accessory_elements = [element for element in accessory_models_collector if "Metal Sleeve" in element.Name or "Plastic Sleeve" in element.Name or "Cast Iron Sleeve" in element.Name]
            for x in accessory_elements:
                slvdiameter = "{0:.2f}".format(get_parameter_value_by_name_AsDouble(x, 'Pipe Nominal Diameter') * 12)
                slvlength = "{0:.2f}".format(get_parameter_value_by_name_AsDouble(x, 'Sleeve Length') * 12)
                result_string = 'SLV {0} x {1}'.format(slvdiameter, slvlength)
                set_parameter_by_name(x, 'TS_Point_Description', result_string)
            accessory_elements2 = [element for element in accessory_models_collector if "Pipe Riser" in element.Name]
            for x in accessory_elements2:
                slvdiameter = "{0:.2f}".format(get_parameter_value_by_name_AsDouble(x, 'Diameter') * 12)
                result_string2 = "{0} RISER".format(slvdiameter)
                set_parameter_by_name(x, 'TS_Point_Description', result_string2)
            accessory_elements3 = [element for element in accessory_models_collector if "Floor Sleeve" in element.Name or "Round Floor Sleeve" in element.Name]
            for x in accessory_elements3:
                slvdiameter = "{0:.2f}".format(get_parameter_value_by_name_AsDouble(x, 'Diameter') * 12)
                slvlength = "{0:.2f}".format(get_parameter_value_by_name_AsDouble(x, 'Length') * 12)
                result_string = 'SLV {0} x {1}'.format(slvdiameter, slvlength)
                set_parameter_by_name(x, 'TS_Point_Description', result_string)
            accessory_elements4 = [element for element in accessory_models_collector if "Rectangular Sleeve" in element.Name]
            for x in accessory_elements4:
                slvlength = "{0:.2f}".format(get_parameter_value_by_name_AsDouble(x, 'Length') * 12)
                slvwidth = "{0:.2f}".format(get_parameter_value_by_name_AsDouble(x, 'Width') * 12)
                slvheight = "{0:.2f}".format(get_parameter_value_by_name_AsDouble(x, 'Height') * 12)
                result_string = 'SLV {0} x {1} x {2}'.format(slvlength, slvwidth, slvheight)
                set_parameter_by_name(x, 'TS_Point_Description', result_string)
            accessory_elements5 = [element for element in accessory_models_collector if "WS" in element.Name or "DR-WS" in element.Name]
            for x in accessory_elements5:
                slvdiameter = "{0:.2f}".format(get_parameter_value_by_name_AsDouble(x, 'Diameter') * 12)
                slvelevation = feet_to_feet_inches_fraction(get_parameter_value_by_name_AsDouble(x, 'Elevation from Level'))
                result_string = "DIA {0} CL {1}".format(slvdiameter, slvelevation)
                set_parameter_by_name(x, 'TS_Point_Description', result_string)
            accessory_elements6 = [element for element in accessory_models_collector if "BLOCKOUT" in element.Name]
            for x in accessory_elements6:
                slvlength = "{0:.2f}".format(get_parameter_value_by_name_AsDouble(x, 'Length') * 12)
                slvwidth = "{0:.2f}".format(get_parameter_value_by_name_AsDouble(x, 'Width') * 12)
                slvheight = "{0:.2f}".format(get_parameter_value_by_name_AsDouble(x, 'Height') * 12)
                result_string = 'BO W{0} x H{1} x L{2}'.format(slvwidth, slvheight, slvlength)
                set_parameter_by_name(x, 'TS_Point_Description', result_string)
            accessory_ductmodels_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_DuctAccessory)
            accessory_elements7 = [element for element in accessory_ductmodels_collector if "RWS" in element.Name]
            for x in accessory_elements7:
                slvlength = "{0:.2f}".format(get_parameter_value_by_name_AsDouble(x, 'Length') * 12)
                slvwidth = "{0:.2f}".format(get_parameter_value_by_name_AsDouble(x, 'Width') * 12)
                slvheight = "{0:.2f}".format(get_parameter_value_by_name_AsDouble(x, 'Height') * 12)
                result_string = 'RWS W{0} x H{1} x L{2}'.format(slvwidth, slvheight, slvlength)
                set_parameter_by_name(x, 'TS_Point_Description', result_string)
    except Exception as e:
        t.RollBack()
        TaskDialog.Show("Error", "Something went wrong: {0}".format(str(e)))
    t.Commit()
except Exception as e:
    TaskDialog.Show("Error", "Failed to execute: {0}".format(str(e)))