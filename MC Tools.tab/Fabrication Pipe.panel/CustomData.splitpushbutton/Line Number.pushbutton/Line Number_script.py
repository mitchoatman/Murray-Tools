import os, clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
import System
from System.Windows import Window, Thickness, HorizontalAlignment, VerticalAlignment, WindowStartupLocation
from System.Windows.Controls import Grid, RowDefinition, ColumnDefinition, Label, TextBox, Button, ListBox, StackPanel
from System.Windows.Input import Keyboard, MouseButtonEventArgs
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, BuiltInCategory, ElementId
from Autodesk.Revit.UI import UIApplication, TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from System.Windows.Interop import WindowInteropHelper
from System.Collections.Generic import List
from Parameters.Add_SharedParameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name

Shared_Params()

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = UIApplication(doc.Application)

def natural_key(s):
    import re
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

# WPF dialog
class LineNumberWindow(Window):
    def __init__(self, default_value, line_numbers, revit_window_handle):
        self.Title = "Line Number"
        self.Width = 300
        self.ResizeMode = System.Windows.ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.InitializeComponents(default_value, line_numbers)
        # Set the owner to Revit's main window
        WindowInteropHelper(self).Owner = revit_window_handle

    def InitializeComponents(self, default_value, line_numbers):
        # Create Grid layout
        grid = Grid()
        self.Content = grid

        # Define rows
        row_definitions = [
            RowDefinition(Height=System.Windows.GridLength.Auto),  # Label
            RowDefinition(Height=System.Windows.GridLength.Auto),  # Textbox
            RowDefinition(Height=System.Windows.GridLength.Auto),  # Listbox label
            RowDefinition(Height=System.Windows.GridLength(1, System.Windows.GridUnitType.Star)),  # Listbox
            RowDefinition(Height=System.Windows.GridLength.Auto)   # Buttons
        ]
        for row in row_definitions:
            grid.RowDefinitions.Add(row)

        # Define columns
        column_definitions = [
            ColumnDefinition(Width=System.Windows.GridLength(1, System.Windows.GridUnitType.Star))
        ]
        grid.ColumnDefinitions.Add(column_definitions[0])

        # Calculate listbox height based on line numbers
        item_height = 20
        listbox_height = item_height * min(15, max(7, len(line_numbers))) + 5
        self.Height = listbox_height + 185

        # Add controls
        row_index = 0

        # Label for TextBox
        self.label = Label()
        self.label.Content = "Enter Line Number:"
        self.label.Margin = Thickness(10, 5, 10, 5)
        Grid.SetRow(self.label, row_index)
        grid.Children.Add(self.label)
        row_index += 1

        # TextBox
        self.textbox = TextBox()
        self.textbox.Text = default_value
        self.textbox.Margin = Thickness(10, 0, 10, 5)
        Grid.SetRow(self.textbox, row_index)
        grid.Children.Add(self.textbox)
        row_index += 1

        # Label for ListBox
        self.list_label = Label()
        self.list_label.Content = "Line Numbers in View:"
        self.list_label.Margin = Thickness(10, 0, 10, 5)
        Grid.SetRow(self.list_label, row_index)
        grid.Children.Add(self.list_label)
        row_index += 1

        # ListBox
        self.listbox = ListBox()
        self.listbox.Height = listbox_height
        self.listbox.Margin = Thickness(10, 0, 10, 0)
        for number in line_numbers:
            self.listbox.Items.Add(number)
        self.listbox.SelectionChanged += self.on_listbox_select
        def listbox_double_click(sender, args):
            if isinstance(args, MouseButtonEventArgs) and args.ClickCount == 2:
                self.on_listbox_double_click(sender, args)
        self.listbox.MouseDoubleClick += listbox_double_click
        Grid.SetRow(self.listbox, row_index)
        grid.Children.Add(self.listbox)
        row_index += 1

        # Buttons
        button_panel = StackPanel()
        button_panel.Orientation = System.Windows.Controls.Orientation.Horizontal
        button_panel.HorizontalAlignment = HorizontalAlignment.Center
        button_panel.Margin = Thickness(0, 15, 0, 10)
        Grid.SetRow(button_panel, row_index)
        grid.Children.Add(button_panel)

        self.ok_button = Button()
        self.ok_button.Content = "OK"
        self.ok_button.Width = 75
        self.ok_button.Height = 25
        self.ok_button.Margin = Thickness(5, 0, 5, 0)
        self.ok_button.Click += self.on_ok_click
        button_panel.Children.Add(self.ok_button)

        self.show_button = Button()
        self.show_button.Content = "Show"
        self.show_button.Width = 75
        self.show_button.Height = 25
        self.show_button.Margin = Thickness(5, 0, 5, 0)
        self.show_button.Click += self.on_show_click
        button_panel.Children.Add(self.show_button)

        self.cancel_button = Button()
        self.cancel_button.Content = "Cancel"
        self.cancel_button.Width = 75
        self.cancel_button.Height = 25
        self.cancel_button.Margin = Thickness(5, 0, 5, 0)
        self.cancel_button.Click += self.on_cancel_click
        button_panel.Children.Add(self.cancel_button)

        # Set Enter and Escape keys
        self.KeyDown += self.on_key_down

    def on_listbox_select(self, sender, event):
        selected = self.listbox.SelectedItem
        if selected:
            self.textbox.Text = selected

    def on_listbox_double_click(self, sender, event):
        selected = self.listbox.SelectedItem
        if selected:
            self.textbox.Text = selected

    def on_show_click(self, sender, event):
        selected = self.listbox.SelectedItem
        if not selected:
            TaskDialog.Show("Warning", "Please select a line number from the list.")
            return

        # Collect elements with the selected FP_Line Number
        matching_elements = []
        collector = FilteredElementCollector(doc, doc.ActiveView.Id)
        for elem in collector:
            param = elem.LookupParameter("FP_Line Number")
            if param and param.HasValue and param.AsString() == selected:
                matching_elements.append(elem)

        if not matching_elements:
            TaskDialog.Show("Warning", "No elements found with line number '{selected}' in the active view.")
            return

        # Zoom to matching elements
        element_ids = List[ElementId]()
        for elem in matching_elements:
            element_ids.Add(elem.Id)
        uidoc.Selection.SetElementIds(element_ids); uidoc.ShowElements(element_ids)

    def on_ok_click(self, sender, event):
        self.DialogResult = True
        self.Close()

    def on_cancel_click(self, sender, event):
        self.DialogResult = False
        self.Close()

    def on_key_down(self, sender, event):
        if event.Key == System.Windows.Input.Key.Enter:
            self.DialogResult = True
            self.Close()
        elif event.Key == System.Windows.Input.Key.Escape:
            self.DialogResult = False
            self.Close()

# Get Revit's main window handle
revit_window_handle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle

# Show dialog
form = LineNumberWindow(PrevInput, line_numbers, revit_window_handle)
value = None
if form.ShowDialog() == True:
    value = form.textbox.Text

# Rest of the existing code, with selection logic after dialog
if value:
    selected_ids = []
    try:
        picked_refs = uidoc.Selection.PickObjects(ObjectType.Element, "Please select elements to set Line Number.")
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
        TaskDialog.Show("Error", "Error: {}".format(str(e)))
        if t is not None and t.HasStarted():
            t.RollBack()