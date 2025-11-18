import clr
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
from System.Windows import Window, Thickness, HorizontalAlignment, VerticalAlignment, WindowStartupLocation, ResizeMode, FontWeights
from System.Windows.Controls import Grid, RowDefinition, ColumnDefinition, Label, TextBox, Button
from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.UI.Selection import ObjectType
from System.Windows.Media import Brushes, FontFamily
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

# WPF dialog
class BundleNumberForm(Window):
    def __init__(self, default_value):
        self.Title = "Bundle Number"
        self.Width = 300
        self.Height = 160
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.NoResize

        self.snumber = None

        grid = Grid()
        grid.Margin = Thickness(10)

        for _ in range(3):
            grid.RowDefinitions.Add(RowDefinition())

        # Label
        label = Label()
        label.Content = "Enter Bundle Number:"
        label.FontFamily = FontFamily("Arial")
        label.FontSize = 14
        label.Margin = Thickness(0, 0, 0, 5)
        Grid.SetRow(label, 0)
        grid.Children.Add(label)

        # TextBox
        textbox = TextBox()
        textbox.Text = default_value
        textbox.Margin = Thickness(0, 0, 0, 10)
        textbox.FontFamily = FontFamily("Arial")
        textbox.FontSize = 12.25
        Grid.SetRow(textbox, 1)
        grid.Children.Add(textbox)

        # Buttons
        button_width = 75
        button_container = Grid()
        button_container.HorizontalAlignment = HorizontalAlignment.Center
        Grid.SetRow(button_container, 2)

        ok_button = Button()
        ok_button.Content = "OK"
        ok_button.Width = button_width
        ok_button.Click += self.ok_clicked
        button_container.Children.Add(ok_button)

        grid.Children.Add(button_container)

        self.Content = grid
        textbox.Focus()
        textbox.SelectAll()

    def ok_clicked(self, sender, args):
        self.snumber = self.Content.Children[1].Text
        self.DialogResult = True
        self.Close()

# Show dialog
form = BundleNumberForm(PrevInput)
value = None
if form.ShowDialog() and form.DialogResult:
    value = form.snumber

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