#Imports
import Autodesk
from pyrevit import revit, forms
from Autodesk.Revit.DB import Transaction
import os
from SharedParam.Add_Parameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name
Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

selection = revit.get_selection()

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_BundleNumber.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)
if not os.path.exists(filepath):
    f = open((filepath), 'w')
    f.write('123')
    f.close()

f = open((filepath), 'r')
PrevInput = f.read()
f.close()

#This displays dialog
value = forms.ask_for_string(default=PrevInput, prompt='Enter Bundle Number:', title='Bundle Number')

if value:

    f = open((filepath), 'w')
    f.write(value)
    f.close()

    # Define function to set custom data by custom id
    def set_customdata_by_custid(fabpart, custid, value):
        fabpart.SetPartCustomDataText(custid, value)

try:
    t = Transaction(doc, 'Set Bundle Number')
    # Start Transaction
    t.Start()

    for i in selection:
        isfabpart = i.LookupParameter("Fabrication Service")
        if isfabpart:
            if i.ItemCustomId == 838:
                set_parameter_by_name(i, "FP_Bundle", value)
                set_customdata_by_custid(i, 6, value)

    # End Transaction
    t.Commit()
except:
    pass

