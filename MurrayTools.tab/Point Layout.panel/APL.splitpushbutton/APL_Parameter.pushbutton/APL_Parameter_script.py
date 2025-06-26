#Imports
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, BuiltInCategory, ElementId
from Autodesk.Revit.UI import Selection, TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString, get_parameter_value_by_name_AsDouble
from Parameters.Add_SharedParameters import Shared_Params
import os
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import Form, Label, TextBox, CheckBox, Button, DialogResult, FormBorderStyle, FormStartPosition
from System.Drawing import Point, Size, Font
from System.Drawing import FontStyle

Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

def feet_to_feet_inches_fraction(decimal_feet, precision=0.125):  # 1/8" = 0.125
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
            return "0'-0 %s\"" % fraction_str
        elif whole_inches > 0 and fraction_str:
            return "0'-%d %s\"" % (whole_inches, fraction_str)
        elif whole_inches > 0:
            return "0'-%d\"" % whole_inches
        else:
            return "0'-0\""
    else:
        if whole_inches == 0 and fraction_str:
            return "%d'-0 %s\"" % (feet, fraction_str)
        elif whole_inches > 0 and fraction_str:
            return "%d'-%d %s\"" % (feet, whole_inches, fraction_str)
        elif whole_inches > 0:
            return "%d'-%d\"" % (feet, whole_inches)
        else:
            return "%d'-0\"" % feet

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

    # WinForm Dialog
    class UpdateAPLForm(Form):
        def __init__(self, desc, pre):
            self.Text = "Update APL Information"
            self.Size = Size(400, 220)
            self.FormBorderStyle = FormBorderStyle.FixedSingle
            self.MaximizeBox = False
            self.MinimizeBox = False
            self.StartPosition = FormStartPosition.CenterScreen

            # Point Prefix
            self.label_pre = Label()
            self.label_pre.Text = "Point Prefix:"
            self.label_pre.Location = Point(20, 20)
            self.label_pre.Size = Size(100, 20)
            self.label_pre.Font = Font(self.label_pre.Font, FontStyle.Bold)
            self.Controls.Add(self.label_pre)

            self.textbox_pre = TextBox()
            self.textbox_pre.Text = pre
            self.textbox_pre.Location = Point(130, 20)
            self.textbox_pre.Size = Size(125, 20)
            self.Controls.Add(self.textbox_pre)

            self.checkbox_pre = CheckBox()
            self.checkbox_pre.Text = "[Enable]"
            self.checkbox_pre.Location = Point(275, 20)
            self.checkbox_pre.Size = Size(70, 20)
            self.Controls.Add(self.checkbox_pre)

            # Point Description
            self.label_desc = Label()
            self.label_desc.Text = "Point Description:"
            self.label_desc.Location = Point(20, 50)
            self.label_desc.Size = Size(100, 20)
            self.label_desc.Font = Font(self.label_desc.Font, FontStyle.Bold)
            self.Controls.Add(self.label_desc)

            self.textbox_desc = TextBox()
            self.textbox_desc.Text = desc
            self.textbox_desc.Location = Point(130, 50)
            self.textbox_desc.Size = Size(125, 20)
            self.Controls.Add(self.textbox_desc)

            self.checkbox_desc = CheckBox()
            self.checkbox_desc.Text = "[Enable]"
            self.checkbox_desc.Location = Point(275, 50)
            self.checkbox_desc.Size = Size(70, 20)
            self.Controls.Add(self.checkbox_desc)

            # Cleanup Insert Description
            self.checkbox_cleanup = CheckBox()
            self.checkbox_cleanup.Text = "Re-format Insert Description (view)"
            self.checkbox_cleanup.Location = Point(20, 80)
            self.checkbox_cleanup.Size = Size(300, 20)
            self.Controls.Add(self.checkbox_cleanup)

            # Sleeve Dimensions
            self.checkbox_slv = CheckBox()
            self.checkbox_slv.Text = "Add Size and Length to Sleeve Description (view)"
            self.checkbox_slv.Location = Point(20, 110)
            self.checkbox_slv.Size = Size(300, 20)
            self.Controls.Add(self.checkbox_slv)

            # OK Button
            self.button_ok = Button()
            self.button_ok.Text = "OK"
            self.button_ok.Location = Point(162, 140)
            self.button_ok.Size = Size(75, 30)
            self.button_ok.Click += self.on_ok
            self.Controls.Add(self.button_ok)

            self.values = {}

        def on_ok(self, sender, args):
            self.values = {
                'Desc': self.textbox_desc.Text,
                'Pre': self.textbox_pre.Text,
                'changepre': self.checkbox_pre.Checked,
                'changedesc': self.checkbox_desc.Checked,
                'cleanupins': self.checkbox_cleanup.Checked,
                'writeslvdims': self.checkbox_slv.Checked
            }
            self.DialogResult = DialogResult.OK
            self.Close()

    # Display dialog
    form = UpdateAPLForm(lines[0], lines[1])
    if form.ShowDialog() != DialogResult.OK:
        raise Exception("Dialog cancelled")

    # Convert dialog input into variable
    value = form.values['Desc'].upper()
    value1 = form.values['Pre'].upper()
    chkpre = form.values['changepre']
    chkdesc = form.values['changedesc']
    chkins = form.values['cleanupins']
    chkslv = form.values['writeslvdims']

    # write values to text file for future retrieval
    with open(filepath, 'w') as the_file:
        line1 = value + '\n'
        line2 = value1 + '\n'
        the_file.writelines([line1, line2])

    # Prompt for selection only if description or prefix checkboxes are checked
    selection = []
    if chkdesc or chkpre:
        OBJselection = uidoc.Selection.PickObjects(ObjectType.Element, 'Select Elements or Finish Button')
        selection = [doc.GetElement(elId) for elId in OBJselection]

    t = Transaction(doc, 'Modify Point Data')
    t.Start()

    try:
        if chkdesc:
            for i in selection:
                if value:
                    param_exist_ts = i.LookupParameter("TS_Point_Description")
                    if param_exist_ts:
                        set_parameter_by_name(i, "TS_Point_Description", value)
                    else:
                        set_parameter_by_name(i, "PointDescription", value)
        if chkpre:
            for i in selection:
                if value1:
                    param_exist = i.LookupParameter("PointNumber")
                    if param_exist:
                        set_parameter_by_name(i, "PointNumber", value1)
                    param_exist_0 = i.LookupParameter("GTP_PointNumber_0")
                    if param_exist_0:
                        set_parameter_by_name(i, "GTP_PointNumber_0", value1)
                    param_exist_1 = i.LookupParameter("GTP_PointNumber_1")
                    if param_exist_1:
                        set_parameter_by_name(i, "GTP_PointNumber_1", value1)
                    param_exist_2 = i.LookupParameter("GTP_PointNumber_2")
                    if param_exist_2:
                        set_parameter_by_name(i, "GTP_PointNumber_2", value1)
                    param_exist_3 = i.LookupParameter("GTP_PointNumber_3")
                    if param_exist_3:
                        set_parameter_by_name(i, "GTP_PointNumber_3", value1)
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
                slvdiameter = "{:.2f}".format(get_parameter_value_by_name_AsDouble(x, 'Pipe Nominal Diameter') * 12)
                slvlength = "{:.2f}".format(get_parameter_value_by_name_AsDouble(x, 'Sleeve Length') * 12)
                result_string = 'SLV ' + slvdiameter + ' x ' + slvlength
                set_parameter_by_name(x, 'TS_Point_Description', result_string)
            accessory_elements2 = [element for element in accessory_models_collector if "Pipe Riser" in element.Name]
            for x in accessory_elements2:
                slvdiameter = "{:.2f}".format(get_parameter_value_by_name_AsDouble(x, 'Diameter') * 12)
                result_string2 = slvdiameter + ' RISER'
                set_parameter_by_name(x, 'TS_Point_Description', result_string2)
            accessory_elements3 = [element for element in accessory_models_collector if "Floor Sleeve" in element.Name or "Round Floor Sleeve" in element.Name]
            for x in accessory_elements3:
                slvdiameter = "{:.2f}".format(get_parameter_value_by_name_AsDouble(x, 'Diameter') * 12)
                slvlength = "{:.2f}".format(get_parameter_value_by_name_AsDouble(x, 'Length') * 12)
                result_string = 'SLV ' + slvdiameter + ' x ' + slvlength
                set_parameter_by_name(x, 'TS_Point_Description', result_string)
            accessory_elements4 = [element for element in accessory_models_collector if "Rectangular Sleeve" in element.Name]
            for x in accessory_elements4:
                slvlength = "{:.2f}".format(get_parameter_value_by_name_AsDouble(x, 'Length') * 12)
                slvwidth = "{:.2f}".format(get_parameter_value_by_name_AsDouble(x, 'Width') * 12)
                slvheight = "{:.2f}".format(get_parameter_value_by_name_AsDouble(x, 'Height') * 12)
                result_string = 'SLV ' + slvlength + ' x ' + slvwidth + ' x ' + slvheight
                set_parameter_by_name(x, 'TS_Point_Description', result_string)
            accessory_elements5 = [element for element in accessory_models_collector if "WS" in element.Name or "DR-WS" in element.Name]
            for x in accessory_elements5:
                slvdiameter = "{:.2f}".format(get_parameter_value_by_name_AsDouble(x, 'Diameter') * 12)
                slvelevation = feet_to_feet_inches_fraction(get_parameter_value_by_name_AsDouble(x, 'Elevation from Level'))
                result_string = "DIA {0} CL {1}".format(slvdiameter, slvelevation)
                set_parameter_by_name(x, 'TS_Point_Description', result_string)
    except:
        TaskDialog.Show("Error", "Something did not get data, good luck! Trust but verify...")
    t.Commit()
except:
    pass