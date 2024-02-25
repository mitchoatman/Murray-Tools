
#Imports
import Autodesk
from pyrevit import DB, revit, script, forms
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI import Selection
from Autodesk.Revit.UI.Selection import ObjectType
import os
from Parameters.Get_Set_Params import set_parameter_by_name


doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

OBJselection = uidoc.Selection.PickObjects(ObjectType.Element, 'SELECT')         
selection = [doc.GetElement( elId ) for elId in OBJselection]

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_PointLayout.txt')

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
value = forms.ask_for_string(default=PrevInput, prompt='Enter Point Description:', title='Point Description')

if value:

    f = open((filepath), 'w')
    f.write(value)
    f.close()

#This displays dialog
value1 = forms.ask_for_string(default=PrevInput, prompt='Enter Point Number Prefix:', title='Prefix')


t = Transaction(doc, 'Modify Point Data')
#Start Transaction
t.Start()

for i in selection:
     if value:
        param_exist_ts = i.LookupParameter("TS_Point_Description")
        if param_exist_ts:
            #writes data to line number parameterzr
            set_parameter_by_name(i,"TS_Point_Description", value)

for i in selection:
    if value1:
        param_exist_0 = i.LookupParameter("PointNumber")
        if param_exist_0:
            #writes data to line number parameterzr
            set_parameter_by_name(i,"PointNumber", value1)


        param_exist_0 = i.LookupParameter("GTP_PointNumber_0")
        if param_exist_0:
            #writes data to line number parameterzr
            set_parameter_by_name(i,"GTP_PointNumber_0", value1)

        param_exist_1 = i.LookupParameter("GTP_PointNumber_1")
        if param_exist_1:
            #writes data to line number parameterzr
            set_parameter_by_name(i,"GTP_PointNumber_1", value1)

        param_exist_2 = i.LookupParameter("GTP_PointNumber_2")
        if param_exist_2:
            #writes data to line number parameterzr
            set_parameter_by_name(i,"GTP_PointNumber_2", value1)

        param_exist_3 = i.LookupParameter("GTP_PointNumber_3")
        if param_exist_3:
            #writes data to line number parameterzr
            set_parameter_by_name(i,"GTP_PointNumber_3", value1)

#End Transaction
t.Commit()
