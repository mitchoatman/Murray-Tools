from itertools import compress
from System.Collections.Generic import List
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.DB.Fabrication import FabricationPartCompareType
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
import System
from System.Windows import Window, Thickness
from System.Windows.Controls import Label, TextBox, CheckBox, Button, Grid, RowDefinition, ColumnDefinition, StackPanel, ScrollViewer
from System.Windows.Media import Brushes, FontFamily
from System.Windows import ResizeMode, HorizontalAlignment, GridLength, GridUnitType
import os, sys, ast
from Autodesk.Revit.Exceptions import OperationCanceledException
from Parameters.Get_Set_Params import set_parameter_by_name
from Parameters.Add_SharedParameters import Shared_Params

Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

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

class IgnoreFieldsForm(Window):
    def __init__(self, ignorebools):
        self.selected_fields = ignorebools[:]
        self.field_list = Bool_List
        self.checkboxes = []
        self.check_all_state = False
        self.filepath = ignore_filepath
        self.InitializeComponents()

    def InitializeComponents(self):
        self.Title = "Select Ignore Fields"
        self.Width = 400
        self.Height = 600
        self.MinWidth = self.Width
        self.MinHeight = self.Height
        self.ResizeMode = ResizeMode.CanResize
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen

        grid = Grid()
        grid.Margin = Thickness(5)
        for i in range(4):  # rows for: label, search box, scroll, buttons
            row = GridLength(1, GridUnitType.Star) if i == 2 else GridLength.Auto
            grid.RowDefinitions.Add(System.Windows.Controls.RowDefinition(Height=row))
        grid.ColumnDefinitions.Add(System.Windows.Controls.ColumnDefinition())

        # Row 0 - Label
        self.label = Label(Content="Search and select fields to ignore:")
        self.label.FontFamily = FontFamily("Arial")
        self.label.FontSize = 16
        self.label.Margin = Thickness(0)
        Grid.SetRow(self.label, 0)
        grid.Children.Add(self.label)

        # Row 1 - Search Box
        self.search_box = TextBox(Height=20, FontFamily=FontFamily("Arial"), FontSize=12)
        self.search_box.TextChanged += self.search_changed
        Grid.SetRow(self.search_box, 1)
        grid.Children.Add(self.search_box)

        # Row 2 - Scrollable Checkbox Panel
        self.checkbox_panel = StackPanel(Orientation=System.Windows.Controls.Orientation.Vertical)
        scroll_viewer = ScrollViewer(Content=self.checkbox_panel, VerticalScrollBarVisibility=System.Windows.Controls.ScrollBarVisibility.Auto)
        scroll_viewer.Margin = Thickness(0, 1, 0, 1)
        Grid.SetRow(scroll_viewer, 2)
        grid.Children.Add(scroll_viewer)

        self.update_checkboxes(self.field_list)

        # Row 3 - Button Panel
        button_panel = StackPanel(Orientation=System.Windows.Controls.Orientation.Horizontal, HorizontalAlignment=HorizontalAlignment.Center, Margin=Thickness(0, 10, 0, 10))

        self.select_button = Button(Content="Select", FontFamily=FontFamily("Arial"), FontSize=12, Margin=Thickness(0, 0, 20, 0))
        self.select_button.Click += self.select_clicked
        button_panel.Children.Add(self.select_button)

        self.check_all_button = Button(Content="Check All", FontFamily=FontFamily("Arial"), FontSize=12)
        self.check_all_button.Click += self.check_all_clicked
        button_panel.Children.Add(self.check_all_button)

        Grid.SetRow(button_panel, 3)
        grid.Children.Add(button_panel)

        self.Content = grid

        # Set focus and highlight search box
        self.search_box.Focus()
        self.search_box.SelectAll()

        # Subscribe to Closed event to save selected_fields
        self.Closed += self.on_closed

    def update_checkboxes(self, fields):
        self.checkbox_panel.Children.Clear()
        self.checkboxes = []
        for field in fields:
            checkbox = CheckBox(Content=field)
            checkbox.Tag = field
            checkbox.IsChecked = field in self.selected_fields
            checkbox.Click += self.checkbox_clicked
            self.checkbox_panel.Children.Add(checkbox)
            self.checkboxes.append(checkbox)

    def search_changed(self, sender, args):
        search_text = self.search_box.Text.lower()
        filtered = [f for f in self.field_list if search_text in f.lower()]
        self.update_checkboxes(filtered)

    def check_all_clicked(self, sender, args):
        self.check_all_state = not self.check_all_state
        for cb in self.checkboxes:
            cb.IsChecked = self.check_all_state
        self.selected_fields = [cb.Content for cb in self.checkboxes if cb.IsChecked]
        with open(self.filepath, 'w') as f:
            f.write(str(self.selected_fields))

    def checkbox_clicked(self, sender, args):
        self.selected_fields = [cb.Content for cb in self.checkboxes if cb.IsChecked]
        with open(self.filepath, 'w') as f:
            f.write(str(self.selected_fields))

    def select_clicked(self, sender, args):
        self.selected_fields = [cb.Content for cb in self.checkboxes if cb.IsChecked]
        with open(self.filepath, 'w') as f:
            f.write(str(self.selected_fields))
        self.DialogResult = True
        self.Close()

    def on_closed(self, sender, args):
        with open(self.filepath, 'w') as f:
            f.write(str(self.selected_fields))

class RenumberForm(Window):
    def __init__(self, prefix, start_num, checkboxdef):
        self.Title = 'Renumber Fabrication Parts'
        self.Width = 400
        self.Height = 460
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.InitializeComponents(prefix, start_num, checkboxdef)

    def InitializeComponents(self, prefix, start_num, checkboxdef):
        grid = Grid()
        grid.Margin = Thickness(10)

        # Define rows
        for i in range(11):
            row = RowDefinition()
            row.Height = GridLength.Auto
            grid.RowDefinitions.Add(row)

        # Label for Prefix
        self.label_prefix = Label()
        self.label_prefix.Content = 'Prefix and Separator:'
        self.label_prefix.Foreground = Brushes.Black
        self.label_prefix.Margin = Thickness(0, 0, 0, 5)
        Grid.SetRow(self.label_prefix, 0)
        grid.Children.Add(self.label_prefix)

        # TextBox for Prefix
        self.textbox_prefix = TextBox()
        self.textbox_prefix.Text = prefix
        self.textbox_prefix.Width = 150
        self.textbox_prefix.Height = 20
        self.textbox_prefix.Margin = Thickness(150, 0, 0, 10)
        self.textbox_prefix.HorizontalAlignment = HorizontalAlignment.Left
        Grid.SetRow(self.textbox_prefix, 0)
        grid.Children.Add(self.textbox_prefix)

        # Label for Start Number
        self.label_startnum = Label()
        self.label_startnum.Content = 'Enter Start Number:'
        self.label_startnum.Foreground = Brushes.Black
        self.label_startnum.Margin = Thickness(0, 0, 0, 5)
        Grid.SetRow(self.label_startnum, 1)
        grid.Children.Add(self.label_startnum)

        # TextBox for Start Number
        self.textbox_startnum = TextBox()
        self.textbox_startnum.Text = start_num
        self.textbox_startnum.Width = 150
        self.textbox_startnum.Height = 20
        self.textbox_startnum.Margin = Thickness(150, 0, 0, 10)
        self.textbox_startnum.HorizontalAlignment = HorizontalAlignment.Left
        Grid.SetRow(self.textbox_startnum, 1)
        grid.Children.Add(self.textbox_startnum)

        # Checkbox for Same Number
        self.checkbox_same = CheckBox()
        self.checkbox_same.Content = 'Same Number for Identical Parts'
        self.checkbox_same.IsChecked = checkboxdef
        self.checkbox_same.Foreground = Brushes.Black
        self.checkbox_same.Margin = Thickness(0, 0, 0, 10)
        Grid.SetRow(self.checkbox_same, 2)
        grid.Children.Add(self.checkbox_same)

        # Checkbox for Start Location
        self.checkbox_startloc = CheckBox()
        self.checkbox_startloc.Content = 'Set Numbering Start Location'
        self.checkbox_startloc.IsChecked = False
        self.checkbox_startloc.Foreground = Brushes.Black
        self.checkbox_startloc.Margin = Thickness(0, 0, 0, 10)
        Grid.SetRow(self.checkbox_startloc, 3)
        grid.Children.Add(self.checkbox_startloc)

        # Button for Ignore Fields
        self.button_ignore = Button()
        self.button_ignore.Content = 'Ignore Fields'
        self.button_ignore.Background = Brushes.Red
        self.button_ignore.Foreground = Brushes.Black
        self.button_ignore.Width = 360
        self.button_ignore.Height = 30
        self.button_ignore.Margin = Thickness(0, 0, 0, 10)
        self.button_ignore.Click += self.on_ignore_click
        Grid.SetRow(self.button_ignore, 4)
        grid.Children.Add(self.button_ignore)

        # Button for All Selected
        self.button_all = Button()
        self.button_all.Content = 'No Filter / Select All'
        self.button_all.Background = Brushes.PaleGoldenrod
        self.button_all.Foreground = Brushes.Black
        self.button_all.Width = 360
        self.button_all.Height = 30
        self.button_all.Margin = Thickness(0, 0, 0, 10)
        self.button_all.Click += self.on_all_click
        Grid.SetRow(self.button_all, 5)
        grid.Children.Add(self.button_all)

        # Label for Filters
        self.label_filters = Label()
        self.label_filters.Content = unichr(8595) + ' Filter Selection ' + unichr(8595)
        self.label_filters.Foreground = Brushes.Black
        self.label_filters.Margin = Thickness(0, 0, 0, 10)
        self.label_filters.HorizontalAlignment = HorizontalAlignment.Center
        Grid.SetRow(self.label_filters, 6)
        grid.Children.Add(self.label_filters)

        # Button for Single Hangers
        self.button_single_hangers = Button()
        self.button_single_hangers.Content = 'Single Hangers'
        self.button_single_hangers.Foreground = Brushes.Black
        self.button_single_hangers.Width = 360
        self.button_single_hangers.Height = 30
        self.button_single_hangers.Margin = Thickness(0, 0, 0, 10)
        self.button_single_hangers.Click += self.on_single_hangers_click
        Grid.SetRow(self.button_single_hangers, 7)
        grid.Children.Add(self.button_single_hangers)

        # Button for Trapeze Hangers
        self.button_trapeze_hangers = Button()
        self.button_trapeze_hangers.Content = 'Trapeze Hangers'
        self.button_trapeze_hangers.Foreground = Brushes.Black
        self.button_trapeze_hangers.Width = 360
        self.button_trapeze_hangers.Height = 30
        self.button_trapeze_hangers.Margin = Thickness(0, 0, 0, 10)
        self.button_trapeze_hangers.Click += self.on_trapeze_hangers_click
        Grid.SetRow(self.button_trapeze_hangers, 8)
        grid.Children.Add(self.button_trapeze_hangers)

        # Button for Pipework
        self.button_pipework = Button()
        self.button_pipework.Content = 'Pipework'
        self.button_pipework.Foreground = Brushes.Black
        self.button_pipework.Width = 360
        self.button_pipework.Height = 30
        self.button_pipework.Margin = Thickness(0, 0, 0, 10)
        self.button_pipework.Click += self.on_pipework_click
        Grid.SetRow(self.button_pipework, 9)
        grid.Children.Add(self.button_pipework)

        # Button for Ductwork
        self.button_ductwork = Button()
        self.button_ductwork.Content = 'Ductwork'
        self.button_ductwork.Foreground = Brushes.Black
        self.button_ductwork.Width = 360
        self.button_ductwork.Height = 30
        self.button_ductwork.Margin = Thickness(0, 0, 0, 10)
        self.button_ductwork.Click += self.on_ductwork_click
        Grid.SetRow(self.button_ductwork, 10)
        grid.Children.Add(self.button_ductwork)

        self.Content = grid
        self.values = {}

    def on_ignore_click(self, sender, args):
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

        form = IgnoreFieldsForm(ignorebools)
        form.ShowDialog()

    def read_ignore_fields(self):
        global IgnBool, IgnFld
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
        IgnBool = [False] * 28
        indices = [Bool_List.index(item) for item in ignorebools if item in Bool_List]
        for index in indices:
            IgnBool[index] = True
        IgnFld = list(compress(ignoreFields, IgnBool))
        return IgnFld

    def on_single_hangers_click(self, sender, args):
        self.values = {
            'prefix': self.textbox_prefix.Text,
            'StrtNum': self.textbox_startnum.Text,
            'checkboxvalue': self.checkbox_same.IsChecked,
            'SetStartcheckboxvalue': self.checkbox_startloc.IsChecked,
            'category': 'MEP Fabrication Hangers',
            'hanger_type': 'single',
            'ignore_fields': self.read_ignore_fields()
        }
        self.DialogResult = True
        self.Close()

    def on_trapeze_hangers_click(self, sender, args):
        self.values = {
            'prefix': self.textbox_prefix.Text,
            'StrtNum': self.textbox_startnum.Text,
            'checkboxvalue': self.checkbox_same.IsChecked,
            'SetStartcheckboxvalue': self.checkbox_startloc.IsChecked,
            'category': 'MEP Fabrication Hangers',
            'hanger_type': 'trapeze',
            'ignore_fields': self.read_ignore_fields()
        }
        self.DialogResult = True
        self.Close()

    def on_pipework_click(self, sender, args):
        self.values = {
            'prefix': self.textbox_prefix.Text,
            'StrtNum': self.textbox_startnum.Text,
            'checkboxvalue': self.checkbox_same.IsChecked,
            'SetStartcheckboxvalue': self.checkbox_startloc.IsChecked,
            'category': 'MEP Fabrication Pipework',
            'ignore_fields': self.read_ignore_fields()
        }
        self.DialogResult = True
        self.Close()

    def on_ductwork_click(self, sender, args):
        self.values = {
            'prefix': self.textbox_prefix.Text,
            'StrtNum': self.textbox_startnum.Text,
            'checkboxvalue': self.checkbox_same.IsChecked,
            'SetStartcheckboxvalue': self.checkbox_startloc.IsChecked,
            'category': 'MEP Fabrication Ductwork',
            'ignore_fields': self.read_ignore_fields()
        }
        self.DialogResult = True
        self.Close()

    def on_all_click(self, sender, args):
        self.values = {
            'prefix': self.textbox_prefix.Text,
            'StrtNum': self.textbox_startnum.Text,
            'checkboxvalue': self.checkbox_same.IsChecked,
            'SetStartcheckboxvalue': self.checkbox_startloc.IsChecked,
            'category': None,
            'ignore_fields': self.read_ignore_fields()
        }
        self.DialogResult = True
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
if form.ShowDialog() != True:
    import sys
    sys.exit()

valuepre = form.values.get('prefix', lines[0])
value = form.values.get('StrtNum', lines[1])
snfip = form.values.get('checkboxvalue', checkboxdef)
sslfn = form.values.get('SetStartcheckboxvalue', False)
category = form.values.get('category', None)
hanger_type = form.values.get('hanger_type', None)
ignore_fields = form.values.get('ignore_fields', IgnFld)

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
    
    start_number = renumber_by_proximity(selected_part, Fhangers1, valuepre, start_number, Fill_length, snfip, ignore_fields)
else:
    if not snfip:
        for ue in Fhangers1:
            num_to_assign = valuepre + str(start_number).zfill(Fill_length)
            set_parameter_by_name(ue, 'Item Number', str(num_to_assign))
            set_parameter_by_name(ue, 'STRATUS Item Number', str(num_to_assign))
            start_number += 1
    else:
        for e in Fhangers1:
            identical_elements = [n for n in Fhangers2 if e.IsSameAs(n, ignore_fields)]
            if RevitINT > 2025:
                key = tuple(element.Id.Value for element in identical_elements)
            else:
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