from pyrevit import DB, forms
from itertools import compress
from System.Collections.Generic import List
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.DB.Fabrication import FabricationPartCompareType
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import ast, os

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def set_parameter_by_name(element, parameterName, value):
    param = element.LookupParameter(parameterName)
    if param and param.IsReadOnly == False:
        if param.StorageType == DB.StorageType.String:
            param.Set(value)
        elif param.StorageType == DB.StorageType.Integer:
            try:
                param.Set(int(value))
            except ValueError:
                pass
        elif param.StorageType == DB.StorageType.Double:
            try:
                param.Set(float(value))
            except ValueError:
                pass

class CustomISelectionFilter(ISelectionFilter):
    def __init__(self, categories):
        self.categories = categories

    def AllowElement(self, e):
        return e.Category.Name in self.categories

    def AllowReference(self, ref, point):
        return True

# List of all fabrication part category names
fabrication_categories = ["MEP Fabrication Hangers", "MEP Fabrication Pipework", "MEP Fabrication Ductwork"]

# Prompt user to choose between individual or window selection
selection_type = forms.CommandSwitchWindow.show(
    ["Pick Elements Individually", "Window Selection"],
    message="Choose Selection Method"
)

if selection_type == "Pick Elements Individually":
    # Pick individual elements in the order of selection
    Fhangers1 = []
    while True:
        try:
            el = uidoc.Selection.PickObject(ObjectType.Element, CustomISelectionFilter(fabrication_categories), "Select a Fabrication Part (Press ESC to finish)")
            Fhangers1.append(doc.GetElement(el.ElementId))
        except:
            break  # Stop when user presses ESC
else:
    # Window selection - order is not important
    pipesel = uidoc.Selection.PickObjects(ObjectType.Element, CustomISelectionFilter(fabrication_categories), "Select Fabrication Parts by window")
    Fhangers1 = [doc.GetElement(elId) for elId in pipesel]

Fhangers2 = Fhangers1[:]  # Copy the list for identical element comparison

# Initial settings for ignoring fields
IgnBool = [False]*28  # Initialize with False to consider all fields

# Define which fields to ignore
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
    ignorebools = []  # Set to empty list if dialog is canceled

with open(filepath, 'w') as f:
    f.write(str(ignorebools))

indices = [Bool_List.index(item) for item in ignorebools]

for index in indices:
    IgnBool[index] = True

IgnFld = list(compress(ignoreFields, IgnBool))

from rpw.ui.forms import FlexForm, Label, TextBox, Button, CheckBox

filepath = os.path.join(folder_name, 'Ribbon_FabRenumber.txt')
if not os.path.exists(folder_name):
    os.makedirs(folder_name)

if not os.path.exists(filepath):
    with open((filepath), 'w') as the_file:
        line1 = ('pre' + '\n')
        line2 = ('num' + '\n')
        line3 = 'False'
        the_file.writelines([line1, line2, line3])

with open((filepath), 'r') as file:
    lines = file.readlines()
    lines = [line.rstrip() for line in lines]

if len(lines) < 3:
    with open((filepath), 'w') as the_file:
        line1 = ('pre' + '\n')
        line2 = ('num' + '\n')
        line3 = 'False'
        the_file.writelines([line1, line2, line3]) 

with open((filepath), 'r') as file:
    lines = file.readlines()
    lines = [line.rstrip() for line in lines]

checkboxdef = lines[2] == 'True'

components = [
    Label('Prefix and Separator:'),
    TextBox('prefix', lines[0]),
    Label('Enter Start Number:'),
    TextBox('StrtNum', lines[1]),
    CheckBox('checkboxvalue', 'Same Number for Identical Parts', default=checkboxdef),
    CheckBox('SetStartcheckboxvalue', 'Set Numbering Start Location', default=False),
    Button('Ok')
]
form = FlexForm('Renumber Fabrication Parts', components)
form.show()

valuepre = form.values['prefix']
value = form.values['StrtNum']
snfip = form.values['checkboxvalue']
sslfn = form.values['SetStartcheckboxvalue']

start_number = int(value)

Fill_length = len(value)

# Function to calculate distance between two fabrication parts
def distance_between_parts(part1, part2):
    point1 = part1.Origin
    point2 = part2.Origin
    return point1.DistanceTo(point2)

# Function to renumber parts by proximity
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

t = Transaction(doc, 'Re-Number Fabrication Hangers')
t.Start()

unique_elements = {}

if sslfn:
    # Prompt user to select the starting part
    selected_part_ref = uidoc.Selection.PickObject(ObjectType.Element, CustomISelectionFilter(fabrication_categories), "Select Fabrication Part to start numbering from")
    selected_part = doc.GetElement(selected_part_ref.ElementId)
    
    # Renumber parts by proximity
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

with open((filepath), 'w') as the_file:
    line1 = (valuepre + '\n')
    line2 = (str(start_number).zfill(Fill_length) + '\n')
    line3 = str(snfip)
    the_file.writelines([line1, line2, line3])