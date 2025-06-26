from itertools import compress
from System.Collections.Generic import List
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.DB.Fabrication import FabricationPartCompareType
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import Form, Label, TextBox, Button, CheckBox, DialogResult, FormStartPosition, FormBorderStyle
from System.Drawing import Point, Size, Font, FontStyle, Color
import os, sys, ast
from Autodesk.Revit.Exceptions import OperationCanceledException
from Parameters.Get_Set_Params import set_parameter_by_name
from Parameters.Add_SharedParameters import Shared_Params

Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, categories, hanger_type=None):
        self.categories = categories
        self.hanger_type = hanger_type

    def AllowElement(self, e):
        if e.Category.Name not in self.categories:
            return False
        if self.hanger_type and e.Category.Name == "MEP Fabrication Hangers":
            rod_info = e.GetRodInfo()
            if rod_info:
                rod_count = rod_info.RodCount
                if self.hanger_type == 'single' and rod_count >= 2:
                    return False
                if self.hanger_type == 'trapeze' and rod_count < 2:
                    return False
            else:
                return False
        return True

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

# Initialize IgnFld based on saved ignore fields
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

class RenumberForm(Form):
    def __init__(self, prefix, start_num, checkboxdef):
        self.Text = 'Renumber Fabrication Parts'
        self.Size = Size(400, 460)
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.StartPosition = FormStartPosition.CenterScreen

        self.label_prefix = Label()
        self.label_prefix.Text = 'Prefix and Separator:'
        self.label_prefix.ForeColor = Color.Black
        self.label_prefix.Font = Font("Arial", 10, FontStyle.Bold)
        self.label_prefix.Location = Point(20, 20)
        self.label_prefix.AutoSize = True
        self.Controls.Add(self.label_prefix)

        self.textbox_prefix = TextBox()
        self.textbox_prefix.Text = prefix
        self.textbox_prefix.Location = Point(170, 20)
        self.textbox_prefix.Size = Size(150, 20)
        self.textbox_prefix.Name = 'prefix'
        self.Controls.Add(self.textbox_prefix)

        self.label_startnum = Label()
        self.label_startnum.Text = 'Enter Start Number:'
        self.label_startnum.ForeColor = Color.Black
        self.label_startnum.Font = Font("Arial", 10, FontStyle.Bold)
        self.label_startnum.Location = Point(20, 50)
        self.label_startnum.AutoSize = True
        self.Controls.Add(self.label_startnum)

        self.textbox_startnum = TextBox()
        self.textbox_startnum.Text = start_num
        self.textbox_startnum.Location = Point(170, 50)
        self.textbox_startnum.Size = Size(150, 20)
        self.textbox_startnum.Name = 'StrtNum'
        self.Controls.Add(self.textbox_startnum)

        self.checkbox_same = CheckBox()
        self.checkbox_same.Text = 'Same Number for Identical Parts'
        self.checkbox_same.Checked = checkboxdef
        self.checkbox_same.ForeColor = Color.Black
        self.checkbox_same.Font = Font("Arial", 10, FontStyle.Regular)
        self.checkbox_same.Location = Point(20, 80)
        self.checkbox_same.AutoSize = True
        self.checkbox_same.Name = 'checkboxvalue'
        self.Controls.Add(self.checkbox_same)

        self.checkbox_startloc = CheckBox()
        self.checkbox_startloc.Text = 'Set Numbering Start Location'
        self.checkbox_startloc.Checked = False
        self.checkbox_startloc.ForeColor = Color.Black
        self.checkbox_startloc.Font = Font("Arial", 10, FontStyle.Regular)
        self.checkbox_startloc.Location = Point(20, 110)
        self.checkbox_startloc.AutoSize = True
        self.checkbox_startloc.Name = 'SetStartcheckboxvalue'
        self.Controls.Add(self.checkbox_startloc)

        self.button_ignore = Button()
        self.button_ignore.Text = 'Ignore Fields'
        self.button_ignore.ForeColor = Color.Black
        self.button_ignore.BackColor = Color.FromArgb(128, 255, 0, 0)
        self.button_ignore.Font = Font("Arial", 10, FontStyle.Regular)
        self.button_ignore.Location = Point(10, 150)
        self.button_ignore.Size = Size(360, 30)
        self.button_ignore.Click += self.on_ignore_click
        self.Controls.Add(self.button_ignore)

        self.button_all = Button()
        self.button_all.Text = 'All Selected'
        self.button_all.ForeColor = Color.Black
        self.button_all.BackColor = Color.FromArgb(128, 255, 255, 0)
        self.button_all.Font = Font("Arial", 10, FontStyle.Regular)
        self.button_all.Location = Point(10, 190)
        self.button_all.Size = Size(360, 30)
        self.button_all.Click += self.on_all_click
        self.Controls.Add(self.button_all)

        self.label_filters = Label()
        self.label_filters.Text = unichr(8595) + ' Filter Selection ' + unichr(8595)
        self.label_filters.ForeColor = Color.Black
        self.label_filters.Font = Font("Arial", 10, FontStyle.Bold)
        self.label_filters.Location = Point(130, 230)
        self.label_filters.AutoSize = True
        self.Controls.Add(self.label_filters)

        self.button_single_hangers = Button()
        self.button_single_hangers.Text = 'Single Hangers'
        self.button_single_hangers.ForeColor = Color.Black
        self.button_single_hangers.Font = Font("Arial", 10, FontStyle.Regular)
        self.button_single_hangers.Location = Point(10, 260)
        self.button_single_hangers.Size = Size(360, 30)
        self.button_single_hangers.Click += self.on_single_hangers_click
        self.Controls.Add(self.button_single_hangers)

        self.button_trapeze_hangers = Button()
        self.button_trapeze_hangers.Text = 'Trapeze Hangers'
        self.button_trapeze_hangers.ForeColor = Color.Black
        self.button_trapeze_hangers.Font = Font("Arial", 10, FontStyle.Regular)
        self.button_trapeze_hangers.Location = Point(10, 300)
        self.button_trapeze_hangers.Size = Size(360, 30)
        self.button_trapeze_hangers.Click += self.on_trapeze_hangers_click
        self.Controls.Add(self.button_trapeze_hangers)

        self.button_pipework = Button()
        self.button_pipework.Text = 'Pipework'
        self.button_pipework.ForeColor = Color.Black
        self.button_pipework.Font = Font("Arial", 10, FontStyle.Regular)
        self.button_pipework.Location = Point(10, 340)
        self.button_pipework.Size = Size(360, 30)
        self.button_pipework.Click += self.on_pipework_click
        self.Controls.Add(self.button_pipework)

        self.button_ductwork = Button()
        self.button_ductwork.Text = 'Ductwork'
        self.button_ductwork.ForeColor = Color.Black
        self.button_ductwork.Font = Font("Arial", 10, FontStyle.Regular)
        self.button_ductwork.Location = Point(10, 380)
        self.button_ductwork.Size = Size(360, 30)
        self.button_ductwork.Click += self.on_ductwork_click
        self.Controls.Add(self.button_ductwork)

        self.values = {}

    def on_ignore_click(self, sender, args):
        from pyrevit import forms
        global IgnBool, IgnFld

        class MyOption(forms.TemplateListItem):
            @property
            def name(self):
                return self.item

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

        ops = [MyOption(item, checked=(item in ignorebools)) for item in Bool_List]

        ignorebools = forms.SelectFromList.show(ops, title='IgnoreField Options', multiselect=True, button_name='Select IgnoreField(s)')
        if ignorebools is None:
            ignorebools = []

        with open(filepath, 'w') as f:
            f.write(str(ignorebools))

        IgnBool = [False] * 28
        indices = [Bool_List.index(item) for item in ignorebools if item in Bool_List]
        for index in indices:
            IgnBool[index] = True
        IgnFld = list(compress(ignoreFields, IgnBool))
        self.values['ignorebools'] = ignorebools

    def on_single_hangers_click(self, sender, args):
        self.values = {
            'prefix': self.textbox_prefix.Text,
            'StrtNum': self.textbox_startnum.Text,
            'checkboxvalue': self.checkbox_same.Checked,
            'SetStartcheckboxvalue': self.checkbox_startloc.Checked,
            'category': 'MEP Fabrication Hangers',
            'hanger_type': 'single'
        }
        self.DialogResult = DialogResult.OK
        self.Close()

    def on_trapeze_hangers_click(self, sender, args):
        self.values = {
            'prefix': self.textbox_prefix.Text,
            'StrtNum': self.textbox_startnum.Text,
            'checkboxvalue': self.checkbox_same.Checked,
            'SetStartcheckboxvalue': self.checkbox_startloc.Checked,
            'category': 'MEP Fabrication Hangers',
            'hanger_type': 'trapeze'
        }
        self.DialogResult = DialogResult.OK
        self.Close()

    def on_pipework_click(self, sender, args):
        self.values = {
            'prefix': self.textbox_prefix.Text,
            'StrtNum': self.textbox_startnum.Text,
            'checkboxvalue': self.checkbox_same.Checked,
            'SetStartcheckboxvalue': self.checkbox_startloc.Checked,
            'category': 'MEP Fabrication Pipework'
        }
        self.DialogResult = DialogResult.OK
        self.Close()

    def on_ductwork_click(self, sender, args):
        self.values = {
            'prefix': self.textbox_prefix.Text,
            'StrtNum': self.textbox_startnum.Text,
            'checkboxvalue': self.checkbox_same.Checked,
            'SetStartcheckboxvalue': self.checkbox_startloc.Checked,
            'category': 'MEP Fabrication Ductwork'
        }
        self.DialogResult = DialogResult.OK
        self.Close()

    def on_all_click(self, sender, args):
        self.values = {
            'prefix': self.textbox_prefix.Text,
            'StrtNum': self.textbox_startnum.Text,
            'checkboxvalue': self.checkbox_same.Checked,
            'SetStartcheckboxvalue': self.checkbox_startloc.Checked,
            'category': None
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
sslfn = form.values.get('SetStartcheckboxvalue', False)
category = form.values.get('category', None)
hanger_type = form.values.get('hanger_type', None)

selected_categories = [category] if category else fabrication_categories
try:
    pipesel = uidoc.Selection.PickObjects(ObjectType.Element, CustomISelectionFilter(selected_categories, hanger_type), "Select Fabrication Parts")
except OperationCanceledException:
    sys.exit()

Fhangers1 = [doc.GetElement(elId) for elId in pipesel]
Fhangers2 = Fhangers1[:]

def distance_between_parts(part1, part2):
    point1 = part1.Origin
    point2 = part2.Origin
    return point1.DistanceTo(point2)

def renumber_by_proximity(selected_part, parts_to_renumber, prefix, start_num, fill_length, identical_parts=False, ignore_fields=None):
    unique_elements = {}
    parts_sorted = sorted(parts_to_renumber, key=lambda x: distance_between_parts(selected_part, x))
    
    current_number = start_num
    
    if not identical_parts:
        for part in parts_sorted:
            num_to_assign = prefix + str(current_number).zfill(fill_length)
            set_parameter_by_name(part, 'Item Number', num_to_assign)
            set_parameter_by_name(part, 'STRATUS Item Number', num_to_assign)
            current_number += 1
    else:
        for part in parts_sorted:
            identical_elements = [n for n in parts_to_renumber if part.IsSameAs(n, ignore_fields)]
            key = tuple(element.Id.IntegerValue for element in identical_elements)
            if key in unique_elements:
                num_to_assign = unique_elements[key]
            else:
                num_to_assign = prefix + str(current_number).zfill(fill_length)
                unique_elements[key] = num_to_assign
                current_number += 1
            for element in identical_elements:
                set_parameter_by_name(element, 'Item Number', num_to_assign)
                set_parameter_by_name(element, 'STRATUS Item Number', num_to_assign)
    
    return current_number

start_number = int(value)
Fill_length = len(value)

if category == 'MEP Fabrication Hangers' and hanger_type:
    if hanger_type == 'single':
        Fhangers1 = [e for e in Fhangers1 if e.GetRodInfo().RodCount < 2]
        Fhangers2 = [e for e in Fhangers2 if e.GetRodInfo().RodCount < 2]
    elif hanger_type == 'trapeze':
        Fhangers1 = [e for e in Fhangers1 if e.GetRodInfo().RodCount > 1]
        Fhangers2 = [e for e in Fhangers2 if e.GetRodInfo().RodCount > 1]

if not Fhangers1:
    import sys
    sys.exit()

t = Transaction(doc, 'Re-Number Fabrication Parts')
t.Start()

unique_elements = {}

if sslfn:
    selected_categories = [category] if category else fabrication_categories
    selected_part_ref = uidoc.Selection.PickObject(ObjectType.Element, CustomISelectionFilter(selected_categories, hanger_type), "Select Fabrication Part to start numbering from")
    selected_part = doc.GetElement(selected_part_ref.ElementId)
    
    start_number = renumber_by_proximity(selected_part, Fhangers1, valuepre, start_number, Fill_length, snfip, IgnFld)
else:
    if not snfip:
        for ue in Fhangers1:
            num_to_assign = valuepre + str(start_number).zfill(Fill_length)
            set_parameter_by_name(ue, 'Item Number', str(num_to_assign))
            set_parameter_by_name(ue, 'STRATUS Item Number', str(num_to_assign))
            start_number += 1
    else:
        for e in Fhangers1:
            identical_elements = [n for n in Fhangers2 if e.IsSameAs(n, IgnFld)]
            key = tuple(element.Id.IntegerValue for element in identical_elements)
            if key in unique_elements:
                num_to_assign = unique_elements[key]
            else:
                num_to_assign = valuepre + str(start_number).zfill(Fill_length)
                unique_elements[key] = num_to_assign
                start_number += 1

            for element in identical_elements:
                set_parameter_by_name(element, 'Item Number', num_to_assign)
                set_parameter_by_name(element, 'STRATUS Item Number', num_to_assign)

t.Commit()

with open(filepath, 'w') as the_file:
    line1 = (valuepre + '\n')
    line2 = (str(start_number).zfill(Fill_length) + '\n')
    line3 = str(snfip)
    the_file.writelines([line1, line2, line3])