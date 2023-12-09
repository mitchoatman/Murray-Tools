
#Imports
import Autodesk
from pyrevit import DB, revit, script, forms
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI import Selection
import os
from SharedParam.Add_Parameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name
Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

selection = revit.get_selection()

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_REFBSDesignation.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)
if not os.path.exists(filepath):
    f = open((filepath), 'w')
    f.write('1')
    f.close()

f = open((filepath), 'r')
PrevInput = f.read()
f.close()

#This displays dialog
value = forms.ask_for_string(default=PrevInput, prompt='Enter REF BS Designation:', title='REF BS Designation')

if value:

    f = open((filepath), 'w')
    f.write(value)
    f.close()
    
    try:
        def set_customdata_by_custid(fabpart, custid, value):
            fabpart.SetPartCustomDataText(custid, value)
        #end of defining functions to use

        t = Transaction(doc, 'Set REF Line Number')
        #Start Transaction
        t.Start()

        for i in selection:
            param_exist = i.LookupParameter("FP_REF BS Designation")
            if param_exist:
                #writes data to line number parameterzr
                set_parameter_by_name(i,"FP_REF BS Designation", value)
        #End Transaction
        t.Commit()
    except:
        pass