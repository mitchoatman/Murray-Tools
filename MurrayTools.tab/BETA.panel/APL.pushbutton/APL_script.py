
#Imports
from pyrevit import DB, revit, script
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI import Selection
from Autodesk.Revit.UI.Selection import ObjectType
from Parameters.Get_Set_Params import set_parameter_by_name
from rpw.ui.forms import FlexForm, Label, TextBox, Button
import os

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

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_PointLayout.txt')
try:
    # read text file for stored values and show them in dialog
    with open((filepath), 'r') as file:
        lines = file.readlines()
        lines = [line.rstrip() for line in lines]
except:
    with open((filepath), 'w') as the_file:
        line1 = ('pre' + '\n')
        line2 = ('num' + '\n')
        the_file.writelines([line1, line2]) 

if len(lines) < 2:
    with open((filepath), 'w') as the_file:
        line1 = ('desc' + '\n')
        line2 = ('pre' + '\n')
        the_file.writelines([line1, line2]) 

# read text file for stored values and show them in dialog
with open((filepath), 'r') as file:
    lines = file.readlines()
    lines = [line.rstrip() for line in lines]


# Display dialog
components = [
    Label('Point Prefix:'),
    TextBox('Pre', lines[1]),
    Label('Point Description:'),
    TextBox('Desc', lines[0]),
    Button('Ok')
    ]
form = FlexForm('Renumber Fabrication Parts', components)
form.show()

# Convert dialog input into variable
value = (form.values['Desc']).upper()
value1 = (form.values['Pre']).upper()

# write values to text file for future retrieval
with open((filepath), 'w') as the_file:
    line1 = (value + '\n')
    line2 = (value1 + '\n')
    the_file.writelines([line1, line2])

t = Transaction(doc, 'Modify Point Data')
#Start Transaction
t.Start()

try:
    for i in selection:
         if value:
            param_exist_ts = i.LookupParameter("TS_Point_Description")
            if param_exist_ts:
                #writes data to line number parameterzr
                set_parameter_by_name(i,"TS_Point_Description", value)
            else:
                set_parameter_by_name(i,"PointDescription", value)  

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
except:
    pass
    print 'Something did not get data, good luck!  Trust but verify...'
#End Transaction
t.Commit()
