
#Imports
import Autodesk
from pyrevit import revit
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, FabricationPart, FabricationConfiguration
import os
from SharedParam.Add_Parameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name
from rpw.ui.forms import FlexForm, Label, ComboBox, TextBox, TextBox, Separator, Button, CheckBox
Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
Config = FabricationConfiguration.GetFabricationConfiguration(doc)
selection = revit.get_selection()

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_StratusAssembly.txt')

# if not os.path.exists(folder_name):
    # os.makedirs(folder_name)
# if not os.path.exists(filepath):
    # f = open((filepath), 'w')
    # f.write('L1-A1-CW-01')
    # f.close()

if not os.path.exists(folder_name):
    os.makedirs(folder_name)
if not os.path.exists(filepath):
    with open((filepath), 'w') as the_file:
        line1 = ('L1-A1-CW-01' + '\n')
        line2 = ('L1-A1-HGR-MAP' + '\n')
        the_file.writelines([line1, line2])

# read text file for stored values and show them in dialog
with open((filepath), 'r') as file:
    lines = file.readlines()
    lines = [line.rstrip() for line in lines]

if len(lines) < 2:
    with open((filepath), 'w') as the_file:
        line1 = ('L1-A1-CW-01' + '\n')
        line2 = ('L1-A1-HGR-MAP' + '\n')
        the_file.writelines([line1, line2]) 

# read text file for stored values and show them in dialog
with open((filepath), 'r') as file:
    lines = file.readlines()
    lines = [line.rstrip() for line in lines]



f = open((filepath), 'r')
PrevInput = f.read()
f.close()

#This displays dialog
components = [Label('Enter Spool Name:'),
    TextBox('spoolname', default=lines[0]),
    Label('Enter Spool Map:'),
    TextBox('spoolmap', default=lines[1]),
    Button('Ok')]
form = FlexForm('Stratus Assembly', components)
form.show()

try:
    value = (form.values['spoolname'])
    map_name = (form.values['spoolmap'])
except:
    pass

# #This displays dialog
# value = forms.ask_for_string(default=PrevInput, prompt='Enter Spool name:', title='Stratus Assembly')
try:
    #splits the spoolname
    valuesplit = value.rsplit('-', 1)

    #gets the length of characters for number
    spoolnumlength = (len(valuesplit[-1]))

    test = (valuesplit[-1]).isnumeric()

    if test:
        #converts number from string to integer
        valuenum = int(float(valuesplit[-1]))

        #increments spool number  by 1
        numincrement = valuenum + 1

        #gets the first half of spool name
        firstpart = valuesplit[0]

        #converts spool number back into string and fills in leading zeros
        lastpart = str(numincrement).zfill(spoolnumlength)

        #combines both halfs of spool name
        newspoolname = firstpart + "-" + lastpart

        f = open((filepath), 'w')
        f.write(newspoolname)
        f.close()

    else:
        f = open((filepath), 'w')
        f.write(value)
        f.close()

    t = Transaction(doc, 'Set Assembly Number')
    #Start Transaction
    t.Start()

    for i in selection:
        param_exist = i.LookupParameter("STRATUS Assembly")
        isfabpart = i.LookupParameter("Fabrication Service")
        if isfabpart:
            stat = i.PartStatus
            STName = Config.GetPartStatusDescription(stat)
            #print (stat)
            #print (STName)
            #writes data to Assembly number parameterzr
            set_parameter_by_name(i,"STRATUS Assembly", value)
            set_parameter_by_name(i,"FP_Spool Map", value)
            set_parameter_by_name(i,"STRATUS Status", "Modeled")
            i.SpoolName = value
            i.PartStatus = 1
            i.Pinned = True
        if param_exist:
            set_parameter_by_name(i,"STRATUS Assembly", value)
            set_parameter_by_name(i,"FP_Spool Map", value)
            set_parameter_by_name(i,"STRATUS Status", "Modeled")
            i.Pinned = True
        else:
            print 'Element selected are missing parameters. Contact admin.'
    #End Transaction
    t.Commit()
except:
    pass