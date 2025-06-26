import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
import System
from System.Windows.Forms import Form, Label, TextBox, Button, DialogResult, FormStartPosition, FormBorderStyle, MessageBox, ListBox
from System.Drawing import Point, Size
from Autodesk.Revit.DB import Transaction, FilteredElementCollector
from Autodesk.Revit.UI import Selection
from Autodesk.Revit.UI.Selection import ObjectType
from Parameters.Add_SharedParameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name
import os
import re

Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def natural_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', s)]

# File handling
folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_LineNumber.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)

if not os.path.exists(filepath):
    with open(filepath, 'w') as f:
        f.write('123')

with open(filepath, 'r') as f:
    PrevInput = f.read()

# Collect unique FP_Line Number values from active view
line_numbers = set()
collector = FilteredElementCollector(doc, doc.ActiveView.Id)
for elem in collector:
    param = elem.LookupParameter("FP_Line Number")
    if param and param.HasValue and param.AsString():
        line_numbers.add(param.AsString())
line_numbers = sorted(line_numbers, key=natural_key)

# WinForms dialog
class LineNumberForm(Form):
    def __init__(self, default_value, line_numbers):
        self.Text = "Line Number"
        self.scale_factor = self.get_dpi_scale()
        self.padding = 5
        self.InitializeComponents(default_value, line_numbers)

    def get_dpi_scale(self):
        screen = System.Windows.Forms.Screen.PrimaryScreen
        graphics = self.CreateGraphics()
        dpi_x = graphics.DpiX
        graphics.Dispose()
        return dpi_x / 96.0

    def scale_value(self, value):
        return int(value * self.scale_factor)

    def InitializeComponents(self, default_value, line_numbers):
        self.FormBorderStyle = FormBorderStyle.FixedSingle
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.StartPosition = FormStartPosition.CenterScreen

        self.Width = self.scale_value(300)

        # ListBox item height
        item_height = self.scale_value(20)

        # Tiered dialog height
        num_items = len(line_numbers)
        if num_items <= 10:
            listbox_height = item_height * 7
        elif num_items <= 15:
            listbox_height = item_height * 11
        elif num_items <= 20:
            listbox_height = item_height * 13
        else:
            listbox_height = item_height * 13

        listbox_height += self.scale_value(10)

        # Fixed height for other UI parts
        header_height = self.scale_value(130)
        button_height = self.scale_value(30)
        bottom_margin = self.scale_value(20)

        self.Height = listbox_height + header_height + button_height + bottom_margin

        # Label for TextBox
        self.label = Label()
        self.label.Text = "Enter Line Number:"
        self.label.Location = Point(self.scale_value(20), self.scale_value(10))
        self.label.Size = Size(self.scale_value(260), self.scale_value(20))
        self.Controls.Add(self.label)

        # TextBox
        self.textbox = TextBox()
        self.textbox.Text = default_value
        self.textbox.Location = Point(self.scale_value(20), self.scale_value(31))
        self.textbox.Size = Size(self.scale_value(240), self.scale_value(20))
        self.Controls.Add(self.textbox)

        # Label for ListBox
        self.list_label = Label()
        self.list_label.Text = "Line Numbers in View:"
        self.list_label.Location = Point(self.scale_value(20), self.scale_value(72))
        self.list_label.Size = Size(self.scale_value(260), self.scale_value(20))
        self.Controls.Add(self.list_label)

        # ListBox
        self.listbox = ListBox()
        self.listbox.Location = Point(self.scale_value(20), self.scale_value(95))
        self.listbox.Size = Size(self.scale_value(240), listbox_height)
        for number in line_numbers:
            self.listbox.Items.Add(number)
        self.listbox.SelectedIndexChanged += self.on_listbox_select
        self.Controls.Add(self.listbox)

        # Buttons
        button_width = self.scale_value(75)
        button_y = self.Height - bottom_margin - bottom_margin - bottom_margin - bottom_margin

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

    def on_listbox_select(self, sender, event):
        selected = self.listbox.SelectedItem
        if selected:
            self.textbox.Text = selected

# Show dialog
form = LineNumberForm(PrevInput, line_numbers)
value = None
if form.ShowDialog() == DialogResult.OK:
    value = form.textbox.Text

if value:
    # Get selected elements after OK is clicked
    selected_ids = uidoc.Selection.GetElementIds()
    if not selected_ids:
        # Prompt user to select elements if none are selected
        try:
            picked_refs = uidoc.Selection.PickObjects(ObjectType.Element, "Please select elements to set Line Number.")
            selected_ids = [ref.ElementId for ref in picked_refs]
        except:
            MessageBox.Show("Selection cancelled. No elements selected.", "Error")
            exit()

    if not selected_ids:
        MessageBox.Show("No elements selected. Please select elements and try again.", "Error")
        exit()

    selection = [doc.GetElement(eid) for eid in selected_ids]

    with open(filepath, 'w') as f:
        f.write(value)

    t = None
    try:
        def set_customdata_by_custid(fabpart, custid, value):
            fabpart.SetPartCustomDataText(custid, value)

        t = Transaction(doc, 'Set Line Number')
        t.Start()

        for i in selection:
            param_exist = i.LookupParameter("FP_Line Number")
            if param_exist and not param_exist.IsReadOnly:
                set_parameter_by_name(i, "FP_Line Number", value)
                if i.LookupParameter("Fabrication Service"):
                    set_customdata_by_custid(i, 1, value)

        t.Commit()
    except Exception as e:
        MessageBox.Show("Error: {}".format(str(e)), "Error")
        if t is not None and t.HasStarted():
            t.RollBack()