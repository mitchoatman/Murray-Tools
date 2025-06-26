import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
import System
from System.Windows.Forms import Form, Label, TextBox, Button, DialogResult, FormStartPosition, FormBorderStyle, MessageBox
from System.Drawing import Point, Size
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI.Selection import ObjectType
import os, sys
from Parameters.Add_SharedParameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name
Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

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

# WinForms dialog
class BundleNumberForm(Form):
    def __init__(self, default_value):
        self.Text = "Bundle Number"
        self.scale_factor = self.get_dpi_scale()
        self.padding = 5
        self.InitializeComponents(default_value)

    def get_dpi_scale(self):
        screen = System.Windows.Forms.Screen.PrimaryScreen
        graphics = self.CreateGraphics()
        dpi_x = graphics.DpiX
        graphics.Dispose()
        return dpi_x / 96.0

    def scale_value(self, value):
        return int(value * self.scale_factor)

    def InitializeComponents(self, default_value):
        self.FormBorderStyle = FormBorderStyle.FixedSingle
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.StartPosition = FormStartPosition.CenterScreen

        self.Width = self.scale_value(300)
        self.Height = self.scale_value(160)

        # Label for TextBox
        self.label = Label()
        self.label.Text = "Enter Bundle Number:"
        self.label.Location = Point(self.scale_value(20), self.scale_value(10))
        self.label.Size = Size(self.scale_value(260), self.scale_value(20))
        self.Controls.Add(self.label)

        # TextBox
        self.textbox = TextBox()
        self.textbox.Text = default_value
        self.textbox.Location = Point(self.scale_value(20), self.scale_value(31))
        self.textbox.Size = Size(self.scale_value(240), self.scale_value(20))
        self.Controls.Add(self.textbox)

        # Buttons
        button_width = self.scale_value(75)
        button_height = self.scale_value(30)
        button_y = self.scale_value(80)

        self.ok_button = Button()
        self.ok_button.Text = "OK"
        self.ok_button.Size = Size(button_width, button_height)
        self.ok_button.Location = Point(self.scale_value(70), button_y)
        self.ok_button.DialogResult = DialogResult.OK
        self.Controls.Add(self.ok_button)

        self.cancel_button = Button()
        self.cancel_button.Text = "Cancel"
        self.cancel_button.Size = Size(button_width, button_height)
        self.cancel_button.Location = Point(self.scale_value(155), button_y)
        self.cancel_button.DialogResult = DialogResult.Cancel
        self.Controls.Add(self.cancel_button)

        self.AcceptButton = self.ok_button
        self.CancelButton = self.cancel_button

# Show dialog
form = BundleNumberForm(PrevInput)
value = None
if form.ShowDialog() == DialogResult.OK:
    value = form.textbox.Text

if value:
    # Get selected elements after OK is clicked
    selected_ids = uidoc.Selection.GetElementIds()
    if not selected_ids:
        # Prompt user to select elements if none are selected
        try:
            picked_refs = uidoc.Selection.PickObjects(ObjectType.Element, "Please select elements to set Bundle Number.")
            selected_ids = [ref.ElementId for ref in picked_refs]
        except:
            MessageBox.Show("Selection cancelled. No elements selected.", "Error")
            sys.exit()

    if not selected_ids:
        MessageBox.Show("No elements selected. Please select elements and try again.", "Error")
        exit()

    selection = [doc.GetElement(eid) for eid in selected_ids]

    f = open((filepath), 'w')
    f.write(value)
    f.close()

    # Define function to set custom data by custom id
    def set_customdata_by_custid(fabpart, custid, value):
        fabpart.SetPartCustomDataText(custid, value)

    t = None
    try:
        t = Transaction(doc, 'Set Bundle Number')
        t.Start()

        for i in selection:
            isfabpart = i.LookupParameter("Fabrication Service")
            if isfabpart:
                if i.ItemCustomId == 838:
                    set_parameter_by_name(i, "FP_Bundle", value)
                    set_customdata_by_custid(i, 6, value)

        t.Commit()
    except Exception as e:
        MessageBox.Show("Error: {}".format(str(e)), "Error")
        if t is not None and t.HasStarted():
            t.RollBack()