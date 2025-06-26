import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
import System
from System.Windows.Forms import Form, Label, TextBox, Button, DialogResult, FormStartPosition, FormBorderStyle
from System.Drawing import Point, Size
import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, TransactionGroup, FabricationPart, FabricationConfiguration
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
import os
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString, get_parameter_value_by_name_AsInteger, get_parameter_value_by_name_AsValueString
from Parameters.Add_SharedParameters import Shared_Params
Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
fec = FilteredElementCollector
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)
Config = FabricationConfiguration.GetFabricationConfiguration(doc)

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

# WinForms dialog
class HangerDataForm(Form):
    def __init__(self, job_number, map_name):
        self.Text = "Hanger Data"
        self.scale_factor = self.get_dpi_scale()
        self.padding = 5
        self.InitializeComponents(job_number, map_name)

    def get_dpi_scale(self):
        screen = System.Windows.Forms.Screen.PrimaryScreen
        graphics = self.CreateGraphics()
        dpi_x = graphics.DpiX
        graphics.Dispose()
        return dpi_x / 96.0

    def scale_value(self, value):
        return int(value * self.scale_factor)

    def InitializeComponents(self, job_number, map_name):
        self.FormBorderStyle = FormBorderStyle.FixedSingle
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.StartPosition = FormStartPosition.CenterScreen

        self.Width = self.scale_value(300)
        self.Height = self.scale_value(200)

        # Label for Job Number
        self.job_label = Label()
        self.job_label.Text = "Enter Job Number:"
        self.job_label.Location = Point(self.scale_value(20), self.scale_value(10))
        self.job_label.Size = Size(self.scale_value(260), self.scale_value(20))
        self.Controls.Add(self.job_label)

        # TextBox for Job Number
        self.job_textbox = TextBox()
        self.job_textbox.Text = job_number
        self.job_textbox.Location = Point(self.scale_value(20), self.scale_value(31))
        self.job_textbox.Size = Size(self.scale_value(240), self.scale_value(20))
        self.Controls.Add(self.job_textbox)

        # Label for Map Name
        self.map_label = Label()
        self.map_label.Text = "Hanger Map Name:"
        self.map_label.Location = Point(self.scale_value(20), self.scale_value(60))
        self.map_label.Size = Size(self.scale_value(260), self.scale_value(20))
        self.Controls.Add(self.map_label)

        # TextBox for Map Name
        self.map_textbox = TextBox()
        self.map_textbox.Text = map_name
        self.map_textbox.Location = Point(self.scale_value(20), self.scale_value(81))
        self.map_textbox.Size = Size(self.scale_value(240), self.scale_value(20))
        self.Controls.Add(self.map_textbox)

        # OK Button
        button_width = self.scale_value(75)
        button_height = self.scale_value(30)
        button_y = self.scale_value(120)

        self.ok_button = Button()
        self.ok_button.Text = "OK"
        self.ok_button.Size = Size(button_width, button_height)
        self.ok_button.Location = Point(self.scale_value(112), button_y)  # Centered
        self.ok_button.DialogResult = DialogResult.OK
        self.Controls.Add(self.ok_button)

        self.AcceptButton = self.ok_button

# Display dialog
form = HangerDataForm(lines[0], lines[1])
try:
    # Convert dialog input into variable
    if form.ShowDialog() == DialogResult.OK:
        JobNumber = form.job_textbox.Text
        MapName = form.map_textbox.Text
    else:
        TaskDialog.Show("Error", "Selection cancelled.")
        raise SystemExit

    # write values to text file for future retrieval
    with open((filepath), 'w') as the_file:
        line1 = JobNumber + '\n'
        line2 = MapName
        the_file.writelines([line1, line2])

    # Get selected elements after dialog is confirmed
    selected_ids = uidoc.Selection.GetElementIds()
    if not selected_ids:
        # Prompt user to select elements if none are selected
        try:
            picked_refs = uidoc.Selection.PickObjects(ObjectType.Element, "Please select elements to set Hanger Data.")
            selected_ids = [ref.ElementId for ref in picked_refs]
        except:
            TaskDialog.Show("Error", "Selection cancelled. No elements selected.")
            raise SystemExit

    if not selected_ids:
        TaskDialog.Show("Error", "No elements selected. Please select elements and try again.")
        raise SystemExit

    selection = [doc.GetElement(eid) for eid in selected_ids]

    t = Transaction(doc, "Update Hanger Data")
    t.Start()

    custom_data_exception_raised = False  # Initialize the flag

    for x in selection:
        isfabpart = x.LookupParameter("Fabrication Service")
        if isfabpart:
            # Update FP parameters
            set_parameter_by_name(x, 'FP_CID', x.ItemCustomId)
            set_parameter_by_name(x, 'FP_Service Type', Config.GetServiceTypeName(x.ServiceType))
            set_parameter_by_name(x, 'FP_Service Name', get_parameter_value_by_name_AsString(x, 'Fabrication Service Name'))
            set_parameter_by_name(x, 'FP_Service Abbreviation', get_parameter_value_by_name_AsString(x, 'Fabrication Service Abbreviation'))
            if x.ItemCustomId == 838:
                set_parameter_by_name(x, 'FP_Rod Attached', 'Yes') if x.GetRodInfo().IsAttachedToStructure else set_parameter_by_name(x, 'FP_Rod Attached', 'No')
                [set_parameter_by_name(x, 'FP_Rod Size', n.AncillaryWidthOrDiameter) for n in x.GetPartAncillaryUsage() if n.AncillaryWidthOrDiameter > 0]
                ProductEntry = x.LookupParameter('Product Entry')
                if ProductEntry:
                    set_parameter_by_name(x, 'FP_Product Entry', get_parameter_value_by_name_AsString(x, 'Product Entry'))
                # Set spool info and custom data for hangers (CID 838)
                try:
                    elev = get_parameter_value_by_name_AsValueString(x, 'Middle Elevation')
                    set_customdata_by_custid(x, 12, JobNumber)
                    set_customdata_by_custid(x, 4, elev)
                except Exception as e:
                    if not custom_data_exception_raised:
                        print('Custom Data error:', e)
                        custom_data_exception_raised = True
                try:
                    stat = x.PartStatus
                    STName = Config.GetPartStatusDescription(stat)
                    set_parameter_by_name(x, "STRATUS Assembly", MapName)
                    set_parameter_by_name(x, "STRATUS Status", "Modeled")
                    x.SpoolName = MapName
                    x.PartStatus = 1
                except Exception as e:
                    print('Parameter error:', e)
            if x.ItemCustomId != 838:
                set_parameter_by_name(x, 'FP_Centerline Length', x.CenterlineLength)
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

except Exception as e:
    TaskDialog.Show("Error", "Error: {}".format(str(e)))
    if 't' in locals() and t.HasStarted():
        t.RollBack()