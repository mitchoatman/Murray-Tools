import clr
import os, sys
import Autodesk
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI.Selection import ObjectType
from Parameters.Add_SharedParameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name
from Autodesk.Revit.UI import TaskDialog

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
from System.Windows import Application, Window, Thickness, WindowStyle, ResizeMode, WindowStartupLocation, HorizontalAlignment
from System.Windows.Controls import Label, TextBox, Button, Grid, RowDefinition

Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_REFLineNumber.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)
if not os.path.exists(filepath):
    with open(filepath, 'w') as f:
        f.write('123')

with open(filepath, 'r') as f:
    PrevInput = f.read()

# WPF Form (no XAML)
class REFLineNumberForm(Window):
    def __init__(self, default_value):
        self.Title = "REF Line Number"
        self.Width = 300
        self.Height = 160
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.result = None

        grid = Grid()
        grid.Margin = Thickness(10)

        for _ in range(3):
            grid.RowDefinitions.Add(RowDefinition())

        self.Content = grid

        label = Label()
        label.Content = "Enter REF Line Number:"
        label.Margin = Thickness(0, -5, 0, 10)
        Grid.SetRow(label, 0)
        grid.Children.Add(label)

        self.textbox = TextBox()
        self.textbox.Text = default_value
        self.textbox.Margin = Thickness(0, 0, 0, 10)
        Grid.SetRow(self.textbox, 1)
        grid.Children.Add(self.textbox)

        ok_button = Button()
        ok_button.Content = "OK"
        ok_button.Width = 75
        ok_button.Height = 25
        ok_button.HorizontalAlignment = HorizontalAlignment.Center
        ok_button.Click += self.on_ok
        Grid.SetRow(ok_button, 2)
        grid.Children.Add(ok_button)

        self.textbox.Focus()
        self.textbox.SelectAll()

    def on_ok(self, sender, args):
        self.result = self.textbox.Text
        self.DialogResult = True
        self.Close()

# Show dialog
form = REFLineNumberForm(PrevInput)
value = None
if form.ShowDialog() and form.DialogResult:
    value = form.result

if value:
    selected_ids = uidoc.Selection.GetElementIds()
    if not selected_ids:
        try:
            picked_refs = uidoc.Selection.PickObjects(ObjectType.Element, "Please select elements to set REF Line Number.")
            selected_ids = [ref.ElementId for ref in picked_refs]
        except:
            TaskDialog.Show("Error", "Selection cancelled. No elements selected.")
            sys.exit()

    if not selected_ids:
        TaskDialog.Show("Error", "No elements selected. Please select elements and try again.")
        sys.exit()

    selection = [doc.GetElement(eid) for eid in selected_ids]

    with open(filepath, 'w') as f:
        f.write(value)

    t = None
    try:
        t = Transaction(doc, 'Set REF Line Number')
        t.Start()

        for i in selection:
            param_exist = i.LookupParameter("FP_REF Line Number")
            if param_exist:
                set_parameter_by_name(i, "FP_REF Line Number", value)

        t.Commit()
    except Exception as e:
        TaskDialog.Show("Error: {}".format(str(e)), "Error")
        if t is not None and t.HasStarted():
            t.RollBack()
