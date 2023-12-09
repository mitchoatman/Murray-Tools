
#Imports
import Autodesk
from pyrevit import revit, DB, script, forms
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, FabricationPart, FabricationConfiguration
import os
from SharedParam.Add_Parameters import Shared_Params

Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
Config = FabricationConfiguration.GetFabricationConfiguration(doc)
selection = revit.get_selection()

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_JobNumber.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)
if not os.path.exists(filepath):
    f = open((filepath), 'w')
    f.write('#####')
    f.close()

f = open((filepath), 'r')
PrevInput = f.read()
f.close()

#This writes to fab part custom data field
def set_customdata_by_custid(fabpart, custid, value):
	fabpart.SetPartCustomDataText(custid, value)
#start of defining functions to use
def set_parameter_by_name(element, parameterName, value):
    element.LookupParameter(parameterName).Set(value)
def get_parameter_value_by_name(element, parameterName):
    return element.LookupParameter(parameterName).AsValueString()

#This displays dialog
value = forms.ask_for_string(default=PrevInput, prompt='Enter Job Number:', title='Job Number')

f = open((filepath), 'w')
f.write(value)
f.close()

t = Transaction(doc, 'Set Line Number')
#Start Transaction
t.Start()

for i in selection:
    isfabpart = i.LookupParameter("Fabrication Service")
    elev = get_parameter_value_by_name(i, 'Middle Elevation')
    if isfabpart:
        set_customdata_by_custid(i, 12, value)
        set_customdata_by_custid(i, 4, elev)
#End Transaction
t.Commit()
