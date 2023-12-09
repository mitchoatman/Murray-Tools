import Autodesk
from pyrevit import revit, DB, script, forms
from itertools import compress
from System.Collections.Generic import List
from Autodesk.Revit.DB import FabricationPart, Transaction
from Autodesk.Revit.DB.Fabrication import *
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

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
  def __init__(self, nom_categorie):
    self.nom_categorie = nom_categorie
  def AllowElement(self, e):
    if e.Category.Name == self.nom_categorie:
      return True
    else:
      return False
  def AllowReference(self, ref, point):
    return True # Corrected capitalization

pipesel = uidoc.Selection.PickObjects(ObjectType.Element, CustomISelectionFilter("MEP Fabrication Hangers"), "Select Fabrication Hangers to Re-Number")
Fhangers = [doc.GetElement(elId) for elId in pipesel]

IgnBool = [False] # Initialize with False to consider all fields

ItmList = []
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

IgnFld = list(compress(ignoreFields, IgnBool))

# This displays dialog
value = forms.ask_for_string(default='1', prompt='Enter Start Number:', title='Number Sequence')

# Converts number from string to integer
start_number = int(value) if value else 1

# Create a dictionary to keep track of the assigned numbers for each hanger
number_mapping = {}
CmpBool = []

# Start a transaction
t = Transaction(doc, 'Re-Number Fabrication Hangers')
t.Start()

# Iterate over the hangers and assign numbers to them, starting with the user-specified number
for hanger in Fhangers:

  # Check if the hanger has already been assigned a number
  if hanger.Id not in number_mapping:

    # Get the next available number
    num_to_assign = start_number
    start_number += 1

    # Assign the number to the hanger
    number_mapping[hanger.Id] = num_to_assign

    # Find all hangers that are identical to the current hanger
    CmpBool = []
    for other_hanger in Fhangers:
      CmpBool.append(hanger.IsSameAs(other_hanger, IgnFld))

    # If the hanger has any identical hangers, assign the same number to all of them
    if any(CmpBool):
      for element in compress(Fhangers, CmpBool):
        set_parameter_by_name(element, 'Item Number', str(num_to_assign))

# Commit the transaction
t.Commit()

if not Fhangers:
  forms.alert('No parts selected.')
