import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
import System
from System.Windows import Window, Thickness
from System.Windows.Controls import Label, TextBox, Button, Grid, RowDefinition, ColumnDefinition
from System.Windows.Media import Brushes
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

# WPF dialog
class HangerDataForm(Window):
    def __init__(self, job_number, map_name):
        self.Title = "Hanger Data"
        self.Width = 300
        self.Height = 220
        self.ResizeMode = System.Windows.ResizeMode.NoResize
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.result = System.Windows.Forms.DialogResult.Cancel
        self.InitializeComponents(job_number, map_name)

    def InitializeComponents(self, job_number, map_name):
        grid = Grid()
        grid.Margin = Thickness(10)
        
        # Define rows
        row1 = RowDefinition()
        row1.Height = System.Windows.GridLength.Auto
        row2 = RowDefinition()
        row2.Height = System.Windows.GridLength.Auto
        row3 = RowDefinition()
        row3.Height = System.Windows.GridLength.Auto
        row4 = RowDefinition()
        row4.Height = System.Windows.GridLength.Auto
        row5 = RowDefinition()
        row5.Height = System.Windows.GridLength.Auto
        grid.RowDefinitions.Add(row1)
        grid.RowDefinitions.Add(row2)
        grid.RowDefinitions.Add(row3)
        grid.RowDefinitions.Add(row4)
        grid.RowDefinitions.Add(row5)

        # Label for Job Number
        self.job_label = Label()
        self.job_label.Content = "Enter Job Number:"
        self.job_label.Margin = Thickness(0, 0, 0, 5)
        Grid.SetRow(self.job_label, 0)
        grid.Children.Add(self.job_label)

        # TextBox for Job Number
        self.job_textbox = TextBox()
        self.job_textbox.Text = job_number
        self.job_textbox.Margin = Thickness(0, 0, 0, 10)
        self.job_textbox.Height = 20
        Grid.SetRow(self.job_textbox, 1)
        grid.Children.Add(self.job_textbox)

        # Label for Map Name
        self.map_label = Label()
        self.map_label.Content = "Hanger Map Name:"
        self.map_label.Margin = Thickness(0, 0, 0, 5)
        Grid.SetRow(self.map_label, 2)
        grid.Children.Add(self.map_label)

        # TextBox for Map Name
        self.map_textbox = TextBox()
        self.map_textbox.Text = map_name
        self.map_textbox.Margin = Thickness(0, 0, 0, 10)
        self.map_textbox.Height = 20
        Grid.SetRow(self.map_textbox, 3)
        grid.Children.Add(self.map_textbox)

        # OK Button
        self.ok_button = Button()
        self.ok_button.Content = "OK"
        self.ok_button.Width = 75
        self.ok_button.Height = 30
        self.ok_button.HorizontalAlignment = System.Windows.HorizontalAlignment.Center
        self.ok_button.Click += self.ok_button_click
        Grid.SetRow(self.ok_button, 4)
        grid.Children.Add(self.ok_button)

        self.Content = grid

        # Set focus and highlight all text in job number textbox
        self.job_textbox.Focus()
        self.job_textbox.SelectAll()

    def ok_button_click(self, sender, args):
        self.result = System.Windows.Forms.DialogResult.OK
        self.Close()

# Display dialog
form = HangerDataForm(lines[0], lines[1])
form.ShowDialog()
try:
    # Convert dialog input into variable
    if form.result == System.Windows.Forms.DialogResult.OK:
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