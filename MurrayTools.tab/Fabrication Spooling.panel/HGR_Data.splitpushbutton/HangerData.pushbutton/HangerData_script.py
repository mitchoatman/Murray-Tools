
#Imports
import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, FabricationPart, FabricationConfiguration
from rpw.ui.forms import FlexForm, Label, ComboBox, TextBox, Separator, Button, CheckBox
from pyrevit import revit
import os
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString, get_parameter_value_by_name_AsInteger, get_parameter_value_by_name_AsValueString
from SharedParam.Add_Parameters import Shared_Params
Shared_Params()

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
fec = FilteredElementCollector
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float (RevitVersion)
Config = FabricationConfiguration.GetFabricationConfiguration(doc)

selection = revit.get_selection()

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_HangerData.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)
if not os.path.exists(filepath):
    with open((filepath), 'w') as the_file:
        line1 = 'Job' + '\n'
        line2 = 'Map' + '\n'
        the_file.writelines([line1, line2])  

# read text file for stored values and show them in dialog
with open((filepath), 'r') as file:
    lines = file.readlines()
    lines = [line.rstrip() for line in lines]

#This writes to fab part custom data field
def set_customdata_by_custid(fabpart, custid, value):
	fabpart.SetPartCustomDataText(custid, value)

# --------------SELECTED ELEMENTS-------------------

selectedelements = []
selectedhangers = []

selection = [doc.GetElement(id) for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]

if selection:
    t = Transaction(doc, "Update FP Parameters")
    t.Start()
    for x in selection:
        isfabpart = x.LookupParameter("Fabrication Service")
        if isfabpart:
            if x.ItemCustomId == 838:
                set_parameter_by_name(x, 'FP_CID', get_parameter_value_by_name_AsInteger(x, 'Part Pattern Number'))
                set_parameter_by_name(x, 'FP_Service Type', Config.GetServiceTypeName(x.ServiceType))
                set_parameter_by_name(x, 'FP_Service Name', get_parameter_value_by_name_AsString(x, 'Fabrication Service Name'))
                set_parameter_by_name(x, 'FP_Service Abbreviation', get_parameter_value_by_name_AsString(x, 'Fabrication Service Abbreviation'))
                set_parameter_by_name(x, 'FP_Hanger Diameter', get_parameter_value_by_name_AsString(x, 'Size of Primary End'))
                set_parameter_by_name(x, 'FP_Rod Attached', 'Yes') if x.GetRodInfo().IsAttachedToStructure else set_parameter_by_name(x, 'FP_Rod Attached', 'No')
                [set_parameter_by_name(x, 'FP_Rod Size', n.AncillaryWidthOrDiameter) for n in x.GetPartAncillaryUsage() if n.AncillaryWidthOrDiameter > 0]
            try:
                if (x.GetRodInfo().RodCount) > 0:
                    ItmDims = x.GetDimensions()
                    for dta in ItmDims:
                        if dta.Name == 'Rod Extn Below':
                            RB = x.GetDimensionValue(dta)
                        if dta.Name == 'Rod Length':
                            RL = x.GetDimensionValue(dta)
                    TRL = RL + RB
                    set_parameter_by_name(x, 'FP_Rod Length', TRL)
            except:
                pass

            try:
                if (x.GetRodInfo().RodCount) > 1:
                    ItmDims = x.GetDimensions()
                    for dta in ItmDims:
                        if dta.Name == 'Width':
                            TrapWidth = x.GetDimensionValue(dta)
                        if dta.Name == 'Bearer Extn':
                            TrapExtn = x.GetDimensionValue(dta)
                        if dta.Name == 'Right Rod Offset':
                            TrapRRod = x.GetDimensionValue(dta)
                        if dta.Name == 'Left Rod Offset':
                            TrapLRod = x.GetDimensionValue(dta)
                    BearerLength = TrapWidth + TrapExtn + TrapExtn
                    set_parameter_by_name(x, 'FP_Bearer Length', BearerLength)
            except:
                pass
    t.Commit()

# ---------------------------------

# Display dialog
components = [
    Label('Enter Job Number:'),
    TextBox('JobNumber', lines[0]),
    Label('Hanger Map Name:'),
    TextBox('MapName', lines[1]),
    Button('Ok')
    ]
form = FlexForm('Hanger Data', components)
form.show()

# Convert dialog input into variable
JobNumber = form.values['JobNumber']
MapName= form.values['MapName']

# write values to text file for future retrieval
with open((filepath), 'w') as the_file:
    line1 = JobNumber + '\n'
    line2 = MapName
    the_file.writelines([line1, line2])


t = Transaction(doc, 'Set Spool Info')
# Start Transaction
t.Start()

for i in selection:
    isfabpart = i.LookupParameter("Fabrication Service")
    if isfabpart:
        if i.ItemCustomId == 838:
            elev = get_parameter_value_by_name_AsValueString(i, 'Middle Elevation')
            try:
                set_customdata_by_custid(i, 12, JobNumber)
                set_customdata_by_custid(i, 4, elev)
            except:
                print 'Database custom data is not correct, contact your admin.'
                pass
            try:
                stat = i.PartStatus
                STName = Config.GetPartStatusDescription(stat)
                set_parameter_by_name(i, "STRATUS Assembly", MapName)
                set_parameter_by_name(i, "STRATUS Status", "Modeled")
                i.SpoolName = MapName
                i.PartStatus = 1
            except:
                print 'Parameters missing, contact your admin.'
                pass
# End Transaction
t.Commit()


