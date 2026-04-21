from itertools import compress
from System.Collections.Generic import List
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.DB.Fabrication import FabricationPartCompareType
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import os, sys, ast
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import Form, Label, TextBox, Button, CheckBox, DialogResult, FormStartPosition, FormBorderStyle
from System.Drawing import Point, Size, Font, FontStyle, Color
from Parameters.Get_Set_Params import set_parameter_by_name
from Parameters.Add_SharedParameters import Shared_Params

Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, categories):
        self.categories = categories

    def AllowElement(self, e):
        return e.Category.Name in self.categories

    def AllowReference(self, ref, point):
        return True

fabrication_categories = ["MEP Fabrication Hangers", "MEP Fabrication Pipework", "MEP Fabrication Ductwork"]

IgnBool = [False]*28

ignoreFields = List[FabricationPartCompareType]()
ignoreFields.Add(FabricationPartCompareType.CutType)
ignoreFields.Add(FabricationPartCompareType.Material)
ignoreFields.Add(FabricationPartCompareType.Specification)
ignoreFields.Add(FabricationPartCompareType.InsulationSpecification)
ignoreFields.Add(FabricationPartCompareType.MaterialGauge)
ignoreFields.Add(FabricationPartCompareType.DuctFacing)
ignoreFields.Add(FabricationPartCompareType.Insulation)
ignoreFields.Add(FabricationPartCompareType.Notes)
ignoreFields.Add(FabricationPartCompareType.Filename)
ignoreFields.Add(FabricationPartCompareType.Description)
ignoreFields.Add(FabricationPartCompareType.CID)
ignoreFields.Add(FabricationPartCompareType.SkinMaterial)
ignoreFields.Add(FabricationPartCompareType.SkinGauge)
ignoreFields.Add(FabricationPartCompareType.Section)
ignoreFields.Add(FabricationPartCompareType.Status)
ignoreFields.Add(FabricationPartCompareType.Service)
ignoreFields.Add(FabricationPartCompareType.Pallet)
ignoreFields.Add(FabricationPartCompareType.BoxNo)
ignoreFields.Add(FabricationPartCompareType.OrderNo)
ignoreFields.Add(FabricationPartCompareType.Drawing)
ignoreFields.Add(FabricationPartCompareType.Zone)
ignoreFields.Add(FabricationPartCompareType.ETag)
ignoreFields.Add(FabricationPartCompareType.Alt)
ignoreFields.Add(FabricationPartCompareType.Spool)
ignoreFields.Add(FabricationPartCompareType.Alias)
ignoreFields.Add(FabricationPartCompareType.PCFKey)
ignoreFields.Add(FabricationPartCompareType.CustomData)
ignoreFields.Add(FabricationPartCompareType.ButtonAlias)

Bool_List = [
    'CutType', 'Material', 'Specification', 'InsulationSpecification', 'MaterialGauge',
    'DuctFacing', 'Insulation', 'Notes', 'Filename', 'Description', 'CID', 'SkinMaterial',
    'SkinGauge', 'Section', 'Status', 'Service', 'Pallet', 'BoxNo', 'OrderNo', 'Drawing',
    'Zone', 'ETag', 'Alt', 'Spool', 'Alias', 'PCFKey', 'CustomData', 'ButtonAlias'
]

folder_name = "c:\\Temp"
ignore_filepath = os.path.join(folder_name, 'Ribbon_FabRenumberOPS.txt')
if not os.path.exists(folder_name):
    os.makedirs(folder_name)

if os.path.exists(ignore_filepath):
    with open(ignore_filepath, 'r') as f:
        try:
            ignorebools = ast.literal_eval(f.read())
            if not isinstance(ignorebools, list):
                ignorebools = []
        except ValueError:
            ignorebools = []
else:
    ignorebools = []

indices = [Bool_List.index(item) for item in ignorebools if item in Bool_List]
for index in indices:
    IgnBool[index] = True
IgnFld = list(compress(ignoreFields, IgnBool))

def show_ignore_fields_dialog():
    from pyrevit import forms
    folder_name = "c:\\Temp"
    filepath = os.path.join(folder_name, 'Ribbon_FabRenumberOPS.txt')

    if not os.path.exists(folder_name):
        os.makedirs(folder_name)

    if os.path.exists(filepath):
        with open(filepath, 'r') as f:
            try:
                ignorebools = ast.literal_eval(f.read())
                if not isinstance(ignorebools, list):
                    ignorebools = []
            except ValueError:
                ignorebools = []
    else:
        ignorebools = []

    ops = [forms.TemplateListItem(item, checked=(item in ignorebools)) for item in Bool_List]

    ignorebools = forms.SelectFromList.show(ops, title='IgnoreField Options', multiselect=True, button_name='Select IgnoreField(s)')
    if ignorebools is None:
        ignorebools = []

    with open(filepath, 'w') as f:
        f.write(str(ignorebools))

    return ignorebools

class RenumberForm(Form):
    def __init__(self, prefix, start_num, checkboxdef):
        self.Text = 'Renumber Fabrication Parts'
        self.Size = Size(290, 270)
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.StartPosition = FormStartPosition.CenterScreen

        self.label_prefix = Label()
        self.label_prefix.Text = 'Prefix and Separator:'
        self.label_prefix.ForeColor = Color.Black
        self.label_prefix.Font = Font("Arial", 10, FontStyle.Bold)
        self.label_prefix.Location = Point(10, 20)
        self.label_prefix.Size = Size(110, 40)
        self.Controls.Add(self.label_prefix)

        self.textbox_prefix = TextBox()
        self.textbox_prefix.Text = prefix
        self.textbox_prefix.Location = Point(125, 20)
        self.textbox_prefix.Size = Size(125, 40)
        self.textbox_prefix.Name = 'prefix'
        self.Controls.Add(self.textbox_prefix)

        self.label_startnum = Label()
        self.label_startnum.Text = 'Enter Start Number:'
        self.label_startnum.ForeColor = Color.Black
        self.label_startnum.Font = Font("Arial", 10, FontStyle.Bold)
        self.label_startnum.Location = Point(10, 70)
        self.label_startnum.Size = Size(110, 40)
        self.Controls.Add(self.label_startnum)

        self.textbox_startnum = TextBox()
        self.textbox_startnum.Text = start_num
        self.textbox_startnum.Location = Point(125, 70)
        self.textbox_startnum.Size = Size(125, 40)
        self.textbox_startnum.Name = 'StrtNum'
        self.Controls.Add(self.textbox_startnum)

        self.checkbox_same = CheckBox()
        self.checkbox_same.Text = 'Same Number for Identical Parts'
        self.checkbox_same.Checked = checkboxdef
        self.checkbox_same.ForeColor = Color.Black
        self.checkbox_same.Font = Font("Arial", 10, FontStyle.Regular)
        self.checkbox_same.Location = Point(10, 120)
        self.checkbox_same.AutoSize = True
        self.checkbox_same.Name = 'checkboxvalue'
        self.Controls.Add(self.checkbox_same)

        self.button_ignore = Button()
        self.button_ignore.Text = 'Ignore Fields'
        self.button_ignore.ForeColor = Color.Black
        self.button_ignore.Font = Font("Arial", 10, FontStyle.Regular)
        self.button_ignore.Location = Point(10, 150)
        self.button_ignore.Size = Size(250, 30)
        self.button_ignore.Click += self.on_ignore_click
        self.Controls.Add(self.button_ignore)

        self.button_select = Button()
        self.button_select.Text = 'Select Element(s)'
        self.button_select.ForeColor = Color.Black
        self.button_select.Font = Font("Arial", 10, FontStyle.Regular)
        self.button_select.Location = Point(10, 190)
        self.button_select.Size = Size(250, 30)
        self.button_select.Click += self.on_select_click
        self.Controls.Add(self.button_select)

        self.values = {}

    def on_ignore_click(self, sender, args):
        global IgnBool, IgnFld
        ignorebools = show_ignore_fields_dialog()
        IgnBool = [False] * 28
        indices = [Bool_List.index(item) for item in ignorebools if item in Bool_List]
        for index in indices:
            IgnBool[index] = True
        IgnFld = list(compress(ignoreFields, IgnBool))
        self.values['ignorebools'] = ignorebools

    def on_select_click(self, sender, args):
        self.values = {
            'prefix': self.textbox_prefix.Text,
            'StrtNum': self.textbox_startnum.Text,
            'checkboxvalue': self.checkbox_same.Checked
        }
        self.DialogResult = DialogResult.OK
        self.Close()

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_FabRenumber.txt')
if not os.path.exists(folder_name):
    os.makedirs(folder_name)

if not os.path.exists(filepath):
    with open(filepath, 'w') as the_file:
        line1 = ('pre' + '\n')
        line2 = ('num' + '\n')
        line3 = 'False'
        the_file.writelines([line1, line2, line3])

with open(filepath, 'r') as file:
    lines = file.readlines()
    lines = [line.rstrip() for line in lines]

if len(lines) < 3:
    with open(filepath, 'w') as the_file:
        line1 = ('pre' + '\n')
        line2 = ('num' + '\n')
        line3 = 'False'
        the_file.writelines([line1, line2, line3])

with open(filepath, 'r') as file:
    lines = file.readlines()
    lines = [line.rstrip() for line in lines]

checkboxdef = lines[2] == 'True'

form = RenumberForm(lines[0], lines[1], checkboxdef)
if form.ShowDialog() != DialogResult.OK:
    import sys
    sys.exit()

valuepre = form.values.get('prefix', lines[0])
value = form.values.get('StrtNum', lines[1])
snfip = form.values.get('checkboxvalue', checkboxdef)

start_number = int(value)
Fill_length = len(value)

Fhangers1 = []
Fhangers2 = []
unique_elements = {}

t = Transaction(doc, 'Re-Number Fabrication Parts')
t.Start()

while True:
    try:
        el = uidoc.Selection.PickObject(ObjectType.Element, CustomISelectionFilter(fabrication_categories), "Select a Fabrication Part (Press ESC to finish)")
        element = doc.GetElement(el.ElementId)
        Fhangers1.append(element)
        Fhangers2.append(element)

        if not snfip:
            num_to_assign = valuepre + str(start_number).zfill(Fill_length)
            set_parameter_by_name(element, 'Item Number', str(num_to_assign))
            set_parameter_by_name(element, 'STRATUS Item Number', str(num_to_assign))
            start_number += 1
        else:
            # Find identical elements
            identical_elements = [n for n in Fhangers2 if element.IsSameAs(n, IgnFld)]
            # Check if any identical element has a number in unique_elements
            for existing_elem in identical_elements[:-1]:  # Exclude the current element
                existing_key = tuple([existing_elem.Id.IntegerValue])
                if existing_key in unique_elements:
                    num_to_assign = unique_elements[existing_key]
                    break
            else:
                # No existing number found, assign a new one
                num_to_assign = valuepre + str(start_number).zfill(Fill_length)
                start_number += 1
            # Store the number for this element's key
            element_key = tuple([element.Id.IntegerValue])
            unique_elements[element_key] = num_to_assign
            # Assign number to the current element only
            set_parameter_by_name(element, 'Item Number', str(num_to_assign))
            set_parameter_by_name(element, 'STRATUS Item Number', str(num_to_assign))

        t.Commit()
        t.Start()

    except:
        break

t.Commit()

with open(filepath, 'w') as the_file:
    line1 = (valuepre + '\n')
    line2 = (str(start_number).zfill(Fill_length) + '\n')
    line3 = str(snfip)
    the_file.writelines([line1, line2, line3])