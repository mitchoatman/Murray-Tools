
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

pipesel = uidoc.Selection.PickObjects(ObjectType.Element, CustomISelectionFilter(fabrication_categories), "Select Fabrication Parts")
Fhangers1 = [doc.GetElement(elId) for elId in pipesel]
Fhangers2 = [doc.GetElement(elId) for elId in pipesel]


IgnBool = [False]*28 # Initialize with False to consider all fields
# IgnBool[15] = True 

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


Bool_List=[
'CutType',
'Material',
'Specification',
'InsulationSpecification',
'MaterialGauge',
'DuctFacing',
'Insulation',
'Notes',
'Filename',
'Description',
'CID',
'SkinMaterial',
'SkinGauge',
'Section',
'Status',
'Service',
'Pallet',
'BoxNo',
'OrderNo',
'Drawing',
'Zone',
'ETag',
'Alt',
'Spool',
'Alias',
'PCFKey',
'CustomData',
'ButtonAlias']

# 0 _ CutType
# 1 _  Material
# 2 _  Specification
# 3 _  InsulationSpecification
# 4 _  MaterialGauge
# 5 _  DuctFacing
# 6 _  Insulation
# 7 _  Notes
# 8 _  Filename
# 9 _  Description
# 10 _ CID
# 11 _ SkinMaterial
# 12 _ SkinGauge
# 13 _ Section
# 14 _ Status
# 15 _ Service
# 16 _ Pallet
# 17 _ BoxNo
# 18 _ OrderNo
# 19 _ Drawing
# 20 _ Zone
# 21 _ ETag
# 22 _ Alt
# 23 _ Spool
# 24 _ Alias
# 25 _ PCFKey
# 26 _ CustomData
# 27 _ ButtonAlias

class MyOption(forms.TemplateListItem):
    @property
    def name(self):
        return self.item

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_FabRenumberOPS.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)

# If the file exists, read the previously selected items
if os.path.exists(filepath):
    with open(filepath, 'r') as f:
        try:
            ignorebools = ast.literal_eval(f.read())
            # Check if ignorebools is a list
            if not isinstance(ignorebools, list):
                ignorebools = []
        except ValueError:
            ignorebools = []
else:
    ignorebools = []

# Create a list of options, setting 'checked' to True for previously selected items
ops = [MyOption(item, checked=(item in ignorebools)) for item in Bool_List]

# Show the selection dialog
ignorebools = forms.SelectFromList.show(ops, title='IgnoreField Options', multiselect=True, button_name='Select IgnoreField(s)')

# Save the selected items to the file
with open(filepath, 'w') as f:
    f.write(str(ignorebools))

# Get and print indices of selected items
indices = [Bool_List.index(item) for item in ignorebools]
# print("Indices of selected items: ", indices)

# Set the values at the selected indices to True
for index in indices:
    IgnBool[index] = True

IgnFld = list(compress(ignoreFields,IgnBool))

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

# read text file for stored values and show them in dialog
with open((filepath), 'r') as file:
    lines = file.readlines()
    lines = [line.rstrip() for line in lines]

if len(lines) < 3:
    with open((filepath), 'w') as the_file:
        line1 = ('pre' + '\n')
        line2 = ('num' + '\n')
        line3 = 'False'
        the_file.writelines([line1, line2, line3]) 

# read text file for stored values and show them in dialog
with open((filepath), 'r') as file:
    lines = file.readlines()
    lines = [line.rstrip() for line in lines]

if lines[2] == 'False':
    checkboxdef = False
else:
    checkboxdef = True

# Display dialog
components = [
    Label('Prefix and Separator:'),
    TextBox('prefix', lines[0]),
    Label('Enter Start Number:'),
    TextBox('StrtNum', lines[1]),
    CheckBox('checkboxvalue', 'Same Number for Identical Parts', default=checkboxdef),
    CheckBox('SetStartcheckboxvalue', 'Set Numbering Start Location (In Development)', default=False),
    Button('Ok')
    ]
form = FlexForm('Renumber Fabrication Parts', components)
form.show()

# Convert dialog input into variable
valuepre = (form.values['prefix'])
value = (form.values['StrtNum'])
snfip = (form.values['checkboxvalue'])
sslfn = (form.values['SetStartcheckboxvalue'])


start_number = int(value)  # Extract the number part starting from the first numeric character

#gets the length of characters for number
Fill_length = (len(value))

# Start a transaction
t = Transaction(doc, 'Re-Number Fabrication Hangers')
t.Start()

# Create a dictionary to keep track of the assigned numbers for each unique element
unique_elements = {}

if snfip == False:
    for ue in Fhangers1:
        num_to_assign = valuepre + str(start_number).zfill(Fill_length)
        set_parameter_by_name(ue, 'Item Number', str(num_to_assign))
        set_parameter_by_name(ue, 'STRATUS Item Number', str(num_to_assign))
        start_number += 1  # Increment the number for the next element
    if sslfn:
        print 'Setting start location not available yet'
else:
    for e in Fhangers1:
        # Create a list to store elements that are identical to 'e'
        identical_elements = [n for n in Fhangers2 if e.IsSameAs(n, IgnFld)]

        # Create a key that uniquely identifies the group of identical elements
        # This could be a combination of several properties of the elements
        key = tuple(element.Id.IntegerValue for element in identical_elements)

        # Check if we have already assigned a number to an identical element
        if key in unique_elements:
            # If so, use the same number
            num_to_assign = unique_elements[key]
        else:
            # If not, assign the next available number and store it in the dictionary
            num_to_assign = valuepre + str(start_number).zfill(Fill_length)
            unique_elements[key] = num_to_assign
            start_number += 1  # Increment the start number for the next unique element

        # Assign the number to all identical elements
        for element in identical_elements:
            set_parameter_by_name(element, 'Item Number', num_to_assign)
            set_parameter_by_name(element, 'STRATUS Item Number', num_to_assign)
    if sslfn:
        print 'Setting start location not available yet'

# Commit the transaction
t.Commit()

# f = open((filepath), 'w')
# f.write (valuepre + str(start_number).zfill(Fill_length))
# f.close()

# write values to text file for future retrieval
with open((filepath), 'w') as the_file:
    line1 = (valuepre + '\n')
    line2 = (str(start_number).zfill(Fill_length) + '\n')
    line3 = str(snfip)
    the_file.writelines([line1, line2, line3])





