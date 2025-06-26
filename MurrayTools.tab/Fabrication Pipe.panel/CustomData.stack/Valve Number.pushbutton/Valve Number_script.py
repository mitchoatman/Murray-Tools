# Imports
import Autodesk
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI import Selection
from Autodesk.Revit.UI.Selection import ObjectType
import os
from Parameters.Add_SharedParameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name
Shared_Params()

# Windows Forms Imports
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import Form, Label, TextBox, Button, DialogResult, FormBorderStyle, FormStartPosition
from System.Drawing import Point, Size

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_ValveNumber.txt')

# Ensure folder and file exist
if not os.path.exists(folder_name):
    os.makedirs(folder_name)
if not os.path.exists(filepath):
    with open(filepath, 'w') as f:
        f.write('123')

# Read previous input
with open(filepath, 'r') as f:
    PrevInput = f.read()

# Create and show Windows Forms dialog
def show_forms_dialog(default_value):
    form = Form()
    form.Text = "Valve Number"
    form.Size = Size(300, 160)
    form.FormBorderStyle = FormBorderStyle.FixedDialog
    form.StartPosition = FormStartPosition.CenterScreen
    form.MaximizeBox = False
    form.MinimizeBox = False

    # Label
    label = Label()
    label.Text = "Enter Valve Number:"
    label.Location = Point(10, 10)
    label.Size = Size(260, 20)
    form.Controls.Add(label)

    # TextBox
    text_box = TextBox()
    text_box.Text = default_value
    text_box.Location = Point(10, 40)
    text_box.Size = Size(260, 25)
    form.Controls.Add(text_box)

    # OK Button
    ok_button = Button()
    ok_button.Text = "OK"
    ok_button.Location = Point(110, 80)
    ok_button.Size = Size(75, 25)
    ok_button.DialogResult = DialogResult.OK
    form.AcceptButton = ok_button
    form.Controls.Add(ok_button)

    # Cancel Button
    cancel_button = Button()
    cancel_button.Text = "Cancel"
    cancel_button.Location = Point(195, 80)
    cancel_button.Size = Size(75, 25)
    cancel_button.DialogResult = DialogResult.Cancel
    form.CancelButton = cancel_button
    form.Controls.Add(cancel_button)

    # Show dialog
    result = form.ShowDialog()

    if result == DialogResult.OK:
        return text_box.Text
    return None

# Display dialog
value = show_forms_dialog(PrevInput)

if value:
    # Get selected elements after OK is clicked
    selected_ids = uidoc.Selection.GetElementIds()
    if not selected_ids:
        # Prompt user to select elements if none are selected
        try:
            picked_refs = uidoc.Selection.PickObjects(ObjectType.Element, "Please select elements to set Valve Number.")
            selected_ids = [ref.ElementId for ref in picked_refs]
        except:
            print("Error: Selection cancelled. No elements selected.")
            raise Exception("Selection cancelled")

    if not selected_ids:
        print("Error: No elements selected. Please select elements and try again.")
        raise Exception("No elements selected")

    selection = [doc.GetElement(eid) for eid in selected_ids]

    # Save new value
    with open(filepath, 'w') as f:
        f.write(value)

    try:
        def set_customdata_by_custid(fabpart, custid, value):
            fabpart.SetPartCustomDataText(custid, value)

        t = Transaction(doc, 'Set Valve Number')
        t.Start()

        for i in selection:
            param_exist = i.LookupParameter("FP_Valve Number")
            if param_exist:
                set_parameter_by_name(i, "FP_Valve Number", value)
                if i.LookupParameter("Fabrication Service"):
                    set_customdata_by_custid(i, 2, value)

        t.Commit()
    except Exception as e:
        print("Error: {}".format(str(e)))
        if t and t.HasStarted():
            t.RollBack()