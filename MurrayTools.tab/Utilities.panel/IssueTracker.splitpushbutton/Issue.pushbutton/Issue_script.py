import clr
clr.AddReference('System')
import System
import System.Diagnostics
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
from System.Windows import Window, Thickness, HorizontalAlignment, WindowStartupLocation, ResizeMode, RoutedEventHandler
from System.Windows.Controls import Grid, RowDefinition, Label, TextBox, Button, StackPanel, Orientation, ComboBox
from System.Windows.Interop import WindowInteropHelper
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, BuiltInCategory, FamilySymbol, Family, ViewType, BuiltInParameter, ElementId, LocationCurve
from Autodesk.Revit.UI import TaskDialog, UIApplication
from Autodesk.Revit.UI.Selection import ObjectType
import os
import sys

# Check active view type
view = __revit__.ActiveUIDocument.ActiveView
if view.ViewType == ViewType.ThreeD:
    TaskDialog.Show("Error", "Cannot use in 3D view.")
    sys.exit()

path, filename = os.path.split(__file__)
NewFilename = '\\MC-ISSUE.rfa'

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
ui_app = __revit__  # UIApplication for MainWindowHandle access

# Get the associated level of the active view
level = doc.ActiveView.GenLevel

# File handling for saving issue type and description
folder_name = "c:\\temp"
filepath = os.path.join(folder_name, 'IssueType.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)

if not os.path.exists(filepath):
    with open(filepath, 'w') as f:
        f.write('ISSUE\n')

with open(filepath, 'r') as f:
    lines = f.readlines()
    prev_type = lines[0].strip() if len(lines) > 0 else 'ISSUE'
    prev_desc = lines[1].strip() if len(lines) > 1 else ''

class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = False
        return True

def select_fabrication_pipe():
    selection = uidoc.Selection
    pipe_ref = selection.PickObject(ObjectType.Element, "Select an MEP Fabrication Pipe")
    pipe = doc.GetElement(pipe_ref.ElementId)
    return pipe

def pick_point():
    picked_point = uidoc.Selection.PickPoint("Pick a point along the centerline of the pipe")
    return picked_point

def get_pipe_centerline(pipe):
    pipe_location = pipe.Location
    if isinstance(pipe_location, LocationCurve):
        return pipe_location.Curve
    else:
        raise Exception("The selected element does not have a valid centerline.")

def project_point_on_curve(point, curve):
    result = curve.Project(point)
    return result.XYZPoint

# WPF modal dialog
class FixtureTypeWindow(Window):
    def __init__(self, revit_window_handle, family_name, document, ui_document, default_type, default_desc):
        self.Title = "Set Issue Type"
        self.Width = 300
        self.Height = 200
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.doc = document
        self.uidoc = ui_document
        self.family_name = family_name
        self.default_type = default_type
        self.default_desc = default_desc
        self.target_symbol = None
        self.description = ""
        self.placement_triggered = False
        self.placement_mode = ""
        self.InitializeComponents()
        self.PopulateComboBox()
        WindowInteropHelper(self).Owner = revit_window_handle

    def InitializeComponents(self):
        # Create Grid layout
        grid = Grid()
        self.Content = grid

        # Define rows
        row_definitions = [
            RowDefinition(Height=System.Windows.GridLength.Auto),  # Issue Type Label
            RowDefinition(Height=System.Windows.GridLength.Auto),  # ComboBox
            RowDefinition(Height=System.Windows.GridLength.Auto),  # Description Label
            RowDefinition(Height=System.Windows.GridLength.Auto),  # TextBox
            RowDefinition(Height=System.Windows.GridLength.Auto)   # Buttons
        ]
        for row in row_definitions:
            grid.RowDefinitions.Add(row)

        # Add controls
        row_index = 0

        # Issue Type Label
        self.issue_label = Label()
        self.issue_label.Content = "Issue Type:"
        self.issue_label.Margin = Thickness(10, 5, 10, 0)
        Grid.SetRow(self.issue_label, row_index)
        grid.Children.Add(self.issue_label)
        row_index += 1

        # ComboBox for Issue Type
        self.combobox = ComboBox()
        self.combobox.IsEditable = True
        self.combobox.Margin = Thickness(10, 0, 10, 5)
        self.combobox.Loaded += self.on_combobox_loaded
        Grid.SetRow(self.combobox, row_index)
        grid.Children.Add(self.combobox)
        row_index += 1

        # Description Label
        self.desc_label = Label()
        self.desc_label.Content = "Description:"
        self.desc_label.Margin = Thickness(10, 5, 10, 0)
        Grid.SetRow(self.desc_label, row_index)
        grid.Children.Add(self.desc_label)
        row_index += 1

        # TextBox for Description
        self.textbox = TextBox()
        self.textbox.Text = self.default_desc
        self.textbox.Margin = Thickness(10, 0, 10, 5)
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
        self.place_button.ToolTip = "Places the family on the floor"
        self.place_button.Margin = Thickness(5, 0, 5, 0)
        self.place_button.Click += self.on_place_click
        button_panel.Children.Add(self.place_button)

        self.place_on_element_button = Button()
        self.place_on_element_button.Content = "Place on Element"
        self.place_on_element_button.Width = 100
        self.place_on_element_button.Height = 25
        self.place_on_element_button.Margin = Thickness(5, 0, 5, 0)
        self.place_on_element_button.Click += self.on_place_on_element_click
        button_panel.Children.Add(self.place_on_element_button)

        self.close_button = Button()
        self.close_button.Content = "Close"
        self.close_button.Width = 75
        self.close_button.Height = 25
        self.close_button.Margin = Thickness(5, 0, 5, 0)
        self.close_button.Click += self.on_close_click
        button_panel.Children.Add(self.close_button)

    def on_combobox_loaded(self, sender, e):
        editable_tb = self.combobox.Template.FindName("PART_EditableTextBox", self.combobox)
        if editable_tb is not None:
            editable_tb.TextChanged += self.on_editable_text_changed

    def on_editable_text_changed(self, sender, e):
        # Convert TextBox input to uppercase
        current_text = sender.Text
        uppercase_text = current_text.upper()
        if current_text != uppercase_text:
            caret_position = sender.CaretIndex
            sender.Text = uppercase_text
            sender.CaretIndex = caret_position

    def PopulateComboBox(self):
        collector = FilteredElementCollector(self.doc).OfCategory(BuiltInCategory.OST_GenericModel).OfClass(FamilySymbol)
        symbols = [sym for sym in collector if sym.Family.Name == self.family_name]
        type_names = [sym.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() for sym in symbols]
        self.combobox.ItemsSource = type_names
        self.combobox.Text = self.default_type

    def find_or_create_symbol(self, issue_type):
        # Search for the family and type
        collector = FilteredElementCollector(self.doc).OfCategory(BuiltInCategory.OST_GenericModel).OfClass(FamilySymbol)
        target_symbol = None
        for symbol in collector:
            if symbol.Family.Name == self.family_name and symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == issue_type:
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
                        new_symbol = base_symbol.Duplicate(issue_type)
                        # Set Description parameter on type to match family type
                        param = new_symbol.LookupParameter("Description")
                        if param and param.StorageType == DB.StorageType.String and not param.IsReadOnly:
                            param.Set(issue_type)
                        elif not param:
                            TaskDialog.Show("Error", "Description parameter not found for family type '{}'.".format(issue_type))
                            t.RollBack()
                            return None
                        elif param.IsReadOnly:
                            TaskDialog.Show("Error", "Description parameter is read-only for family type '{}'.".format(issue_type))
                            t.RollBack()
                            return None
                        target_symbol = new_symbol
                    else:
                        TaskDialog.Show("Error", "No symbols found in family '{}'.".format(self.family_name))
                        t.RollBack()
                        return None
                else:
                    TaskDialog.Show("Error", "Family '{}' not found.".format(self.family_name))
                    t.RollBack()
                    return None
                t.Commit()
            except Exception as e:
                TaskDialog.Show("Error", "Error creating new type '{}': {}".format(issue_type, str(e)))
                if t.HasStarted():
                    t.RollBack()
                return None

        return target_symbol

    def on_place_click(self, sender, event):
        issue_type = self.combobox.Text.strip()
        self.description = self.textbox.Text.strip()
        if not issue_type:
            TaskDialog.Show("Error", "Please enter a valid issue type.")
            return

        target_symbol = self.find_or_create_symbol(issue_type)
        if target_symbol is None:
            self.placement_triggered = False
            return

        self.target_symbol = target_symbol
        self.placement_mode = "direct"
        self.placement_triggered = True
        self.Close()

    def on_place_on_element_click(self, sender, event):
        issue_type = self.combobox.Text.strip()
        self.description = self.textbox.Text.strip()
        if not issue_type:
            TaskDialog.Show("Error", "Please enter a valid issue type.")
            return

        target_symbol = self.find_or_create_symbol(issue_type)
        if target_symbol is None:
            self.placement_triggered = False
            return

        self.target_symbol = target_symbol
        self.placement_mode = "element"
        self.placement_triggered = True
        self.Close()

    def on_close_click(self, sender, event):
        self.placement_triggered = False
        self.Close()

# Search project for all Families
families = FilteredElementCollector(doc).OfClass(Family)
FamilyName = 'MC-ISSUE'
Fam_is_in_project = any(f.Name == FamilyName for f in families)

family_pathCC = path + NewFilename

# Load family if not present
t = Transaction(doc, 'Load MC-ISSUE Family')
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

# Show modal dialog
form = FixtureTypeWindow(revit_window_handle, FamilyName, doc, uidoc, prev_type, prev_desc)
form.ShowDialog()

# Save the current issue type and description
with open(filepath, 'w') as f:
    f.write(form.combobox.Text.strip() + '\n' + form.textbox.Text.strip())

# Proceed with placement if triggered
if form.placement_triggered and form.target_symbol:
    if not form.target_symbol.IsActive:
        t = Transaction(doc, 'Activate Family Symbol')
        t.Start()
        try:
            form.target_symbol.Activate()
            t.Commit()
        except Exception as e:
            TaskDialog.Show("Error", "Error activating symbol: {}".format(str(e)))
            if t.HasStarted():
                t.RollBack()
            sys.exit()
    
    try:
        t = Transaction(doc, 'Place MC-ISSUE Instance')
        t.Start()
        try:
            if form.placement_mode == "direct":
                insertion_point = uidoc.Selection.PickPoint("Pick insertion point for MC-ISSUE")
            elif form.placement_mode == "element":
                pipe = select_fabrication_pipe()
                centerline_curve = get_pipe_centerline(pipe)
                picked_point = pick_point()
                projected_point = project_point_on_curve(picked_point, centerline_curve)
                insertion_point = DB.XYZ(picked_point.X, picked_point.Y, projected_point.Z)
            else:
                raise Exception("Invalid placement mode.")
            new_family_instance = doc.Create.NewFamilyInstance(insertion_point, form.target_symbol, DB.Structure.StructuralType.NonStructural)
            # Set Comments on the placed instance as soon as it is created
            param = new_family_instance.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
            if param and param.StorageType == DB.StorageType.String and not param.IsReadOnly:
                param.Set(form.description)
            # Set Schedule Level parameter to match the active view's level
            schedule_level_param = new_family_instance.LookupParameter("Schedule Level")
            if schedule_level_param:
                schedule_level_param.Set(level.Id)
            t.Commit()
        except Exception as e:
            TaskDialog.Show("Warning", "Could not place instance or set parameters: {}".format(str(e)))
            t.RollBack()
    except Exception as e:
        pass  # Ignore cancellation exceptions