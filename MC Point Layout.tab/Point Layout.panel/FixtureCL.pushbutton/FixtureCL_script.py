import clr
clr.AddReference('System')
import System
import System.Diagnostics
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
from System.Windows import Window, Thickness, HorizontalAlignment, WindowStartupLocation, ResizeMode
from System.Windows.Controls import Grid, RowDefinition, Label, TextBox, Button, StackPanel, Orientation
from System.Windows.Interop import WindowInteropHelper
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, BuiltInCategory, FamilySymbol, Family, ViewType
from Autodesk.Revit.UI import TaskDialog, UIApplication
import os
import sys

# Check active view type
view = __revit__.ActiveUIDocument.ActiveView
if view.ViewType == ViewType.ThreeD:
    TaskDialog.Show("Error", "Cannot use in 3D view.")
    sys.exit()

path, filename = os.path.split(__file__)
NewFilename = '\\FixtureCL.rfa'

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
ui_app = __revit__  # UIApplication for MainWindowHandle access

# File handling for saving fixture type
folder_name = "c:\\temp"
filepath = os.path.join(folder_name, 'FixtureType.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)

if not os.path.exists(filepath):
    with open(filepath, 'w') as f:
        f.write('WC-1')

with open(filepath, 'r') as f:
    prev_input = f.read().strip()

class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True

# WPF modal dialog
class FixtureTypeWindow(Window):
    def __init__(self, revit_window_handle, family_name, document, ui_document, default_input):
        self.Title = "Set Fixture Type"
        self.Width = 300
        self.Height = 150
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.doc = document
        self.uidoc = ui_document
        self.family_name = family_name
        self.default_input = default_input
        self.selected_type = None  # Store selected type for saving
        self.placement_triggered = False  # Track if placement was initiated
        self.InitializeComponents()
        WindowInteropHelper(self).Owner = revit_window_handle

    def InitializeComponents(self):
        # Create Grid layout
        grid = Grid()
        self.Content = grid

        # Define rows
        row_definitions = [
            RowDefinition(Height=System.Windows.GridLength.Auto),  # Label
            RowDefinition(Height=System.Windows.GridLength.Auto),  # TextBox
            RowDefinition(Height=System.Windows.GridLength.Auto)   # Buttons
        ]
        for row in row_definitions:
            grid.RowDefinitions.Add(row)

        # Add controls
        row_index = 0

        # Label for TextBox
        self.label = Label()
        self.label.Content = "Enter Fixture Type (e.g., WC-1, L-1):"
        self.label.Margin = Thickness(10, 5, 10, 5)
        Grid.SetRow(self.label, row_index)
        grid.Children.Add(self.label)
        row_index += 1

        # TextBox with uppercase conversion
        self.textbox = TextBox()
        self.textbox.Text = self.default_input
        self.textbox.Margin = Thickness(10, 0, 10, 5)
        self.textbox.TextChanged += self.on_text_changed
        Grid.SetRow(self.textbox, row_index)
        grid.Children.Add(self.textbox)
        row_index += 1

        # Buttons
        button_panel = StackPanel()
        button_panel.Orientation = Orientation.Horizontal
        button_panel.HorizontalAlignment = HorizontalAlignment.Center
        button_panel.Margin = Thickness(0, 10, 0, 10)
        Grid.SetRow(button_panel, row_index)
        grid.Children.Add(button_panel)

        self.place_button = Button()
        self.place_button.Content = "Place"
        self.place_button.Width = 75
        self.place_button.Height = 25
        self.place_button.Margin = Thickness(5, 0, 5, 0)
        self.place_button.Click += self.on_place_click
        button_panel.Children.Add(self.place_button)

        self.close_button = Button()
        self.close_button.Content = "Close"
        self.close_button.Width = 75
        self.close_button.Height = 25
        self.close_button.Margin = Thickness(5, 0, 5, 0)
        self.close_button.Click += self.on_close_click
        button_panel.Children.Add(self.close_button)

    def on_text_changed(self, sender, event):
        # Convert TextBox input to uppercase
        current_text = self.textbox.Text
        uppercase_text = current_text.upper()
        if current_text != uppercase_text:
            caret_position = self.textbox.CaretIndex
            self.textbox.Text = uppercase_text
            self.textbox.CaretIndex = caret_position

    def on_place_click(self, sender, event):
        fixture_type = self.textbox.Text.strip()
        if not fixture_type:
            TaskDialog.Show("Error", "Please enter a valid fixture type.")
            return

        self.selected_type = fixture_type  # Save for writing to file
        self.placement_triggered = True  # Indicate placement is active
        # Search for the family and type
        collector = FilteredElementCollector(self.doc).OfCategory(BuiltInCategory.OST_PipeAccessory).OfClass(FamilySymbol)
        target_symbol = None
        for symbol in collector:
            if symbol.Family.Name == self.family_name and symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == fixture_type:
                target_symbol = symbol
                break

        # If type not found, create it
        if not target_symbol:
            t = Transaction(self.doc, 'Create New Family Type')
            t.Start()
            try:
                family = None
                families = FilteredElementCollector(self.doc).OfClass(Family)
                for f in families:
                    if f.Name == self.family_name:
                        family = f
                        break
                
                if family:
                    # Get the first symbol to duplicate
                    symbol_ids = family.GetFamilySymbolIds()
                    if symbol_ids.Count > 0:  # Check if HashSet is not empty
                        base_symbol_id = next(iter(symbol_ids))  # Get first ElementId
                        base_symbol = self.doc.GetElement(base_symbol_id)
                        new_symbol = base_symbol.Duplicate(fixture_type)
                        # Set TS_Point_Description parameter
                        param = new_symbol.LookupParameter("TS_Point_Description")
                        if param and param.StorageType == DB.StorageType.String and not param.IsReadOnly:
                            param.Set(fixture_type)
                        elif not param:
                            TaskDialog.Show("Error", "Parameter 'TS_Point_Description' not found in family type '{}'.".format(fixture_type))
                            t.RollBack()
                            self.placement_triggered = False
                            return
                        elif param.IsReadOnly:
                            TaskDialog.Show("Error", "Parameter 'TS_Point_Description' is read-only in family type '{}'.".format(fixture_type))
                            t.RollBack()
                            self.placement_triggered = False
                            return
                        target_symbol = new_symbol
                    else:
                        TaskDialog.Show("Error", "No symbols found in family '{}'.".format(self.family_name))
                        t.RollBack()
                        self.placement_triggered = False
                        return
                else:
                    TaskDialog.Show("Error", "Family '{}' not found.".format(self.family_name))
                    t.RollBack()
                    self.placement_triggered = False
                    return
                t.Commit()
            except Exception as e:
                TaskDialog.Show("Error", "Error creating new type '{}': {}".format(fixture_type, str(e)))
                if t.HasStarted():
                    t.RollBack()
                self.placement_triggered = False
                return

        # Activate and place the symbol
        if target_symbol:
            if not target_symbol.IsActive:
                t = Transaction(self.doc, 'Activate Family Symbol')
                t.Start()
                try:
                    target_symbol.Activate()
                    t.Commit()
                except Exception as e:
                    TaskDialog.Show("Error", "Error activating symbol: {}".format(str(e)))
                    if t.HasStarted():
                        t.RollBack()
                    self.placement_triggered = False
                    return
            
            try:
                self.Close()  # Close dialog to enter placement mode
                self.uidoc.PromptForFamilyInstancePlacement(target_symbol)
            except Exception as e:
                pass  # Ignore cancellation exceptions

    def on_close_click(self, sender, event):
        self.placement_triggered = False  # No placement if closing
        self.Close()

# Search project for all Families
families = FilteredElementCollector(doc).OfClass(Family)
FamilyName = 'FixtureCL'
Fam_is_in_project = any(f.Name == FamilyName for f in families)

family_pathCC = path + NewFilename

# Load family if not present
t = Transaction(doc, 'Load Fixture CL Symbol Family')
t.Start()
if not Fam_is_in_project:
    fload_handler = FamilyLoaderOptionsHandler()
    try:
        doc.LoadFamily(family_pathCC, fload_handler)
    except Exception as e:
        TaskDialog.Show("Error", "Error loading family: {}".format(str(e)))
        t.RollBack()
        sys.exit()
t.Commit()

# Get Revit's main window handle
revit_window_handle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle

# Show modal dialog once
form = FixtureTypeWindow(revit_window_handle, FamilyName, doc, uidoc, prev_input)
form.ShowDialog()
# Save the selected fixture type if placement was attempted
if form.selected_type:
    with open(filepath, 'w') as f:
        f.write(form.selected_type)