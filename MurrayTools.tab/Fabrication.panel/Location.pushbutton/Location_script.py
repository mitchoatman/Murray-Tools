
#Imports
import Autodesk
from pyrevit import DB, revit, forms
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI import Selection
import os, sys
from Parameters.Add_SharedParameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name
Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

selection = revit.get_selection()

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_Location.txt')

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
value = forms.ask_for_string(default=PrevInput, prompt='Enter Location:', title='Location')

if value:

    f = open((filepath), 'w')
    f.write(value)
    f.close()
    
    try:
        def set_customdata_by_custid(fabpart, custid, value):
            fabpart.SetPartCustomDataText(custid, value)
        #end of defining functions to use
    except:
        print "Error:", sys.exc_info()[0]

    t = Transaction(doc, 'Fabrication Location')
    #Start Transaction
    t.Start()
    try:
        for i in selection:
            param_exist = i.LookupParameter("FP_Location")
            if param_exist:
                #writes data to line number parameterzr
                set_parameter_by_name(i,"FP_Location", value)
                #checks if element has fab part parameter, if yes it writes to fab part custom data
                if i.LookupParameter("Fabrication Service"):
                    set_customdata_by_custid(i, 13, value)
        #End Transaction
        t.Commit()
    except:
        print "Error:", sys.exc_info()[0]