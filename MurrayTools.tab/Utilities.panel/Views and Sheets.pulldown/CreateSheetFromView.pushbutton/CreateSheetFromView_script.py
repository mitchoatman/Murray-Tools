# coding: utf8
import Autodesk
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xaml')
import System
from System.Windows.Controls import Label, TextBox, Button, ScrollViewer, StackPanel, Grid, Orientation, CheckBox
from System.Windows import Window, Thickness, SizeToContent, ResizeMode, HorizontalAlignment, VerticalAlignment, GridLength, GridUnitType
from System.Windows.Media import Brushes, FontFamily
from Autodesk.Revit.DB import FilteredElementCollector, BoundingBoxXYZ, XYZ, BuiltInCategory, Transaction, ViewSheet, Viewport, BuiltInParameter
from Autodesk.Revit.UI import TaskDialog, Selection
import sys
from collections import defaultdict

# Define the active Revit application and document
DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
fec = FilteredElementCollector
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

class ViewSelectionFilter(Window):
    def __init__(self, views):
        self.selected_views = []
        self.view_list = sorted(views, key=lambda x: x.Name)
        self.checkboxes = []
        self.check_all_state = False
        self.InitializeComponents()

    def InitializeComponents(self):
        self.Title = "Select Views"
        self.Width = 400
        self.Height = 400
        self.MinWidth = self.Width
        self.MinHeight = self.Height
        self.ResizeMode = ResizeMode.CanResize
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen

        grid = Grid()
        grid.Margin = Thickness(5)
        for i in range(4):  # rows for: label, search box, scroll, buttons
            row = GridLength(1, GridUnitType.Star) if i == 2 else GridLength.Auto
            grid.RowDefinitions.Add(System.Windows.Controls.RowDefinition(Height=row))
        grid.ColumnDefinitions.Add(System.Windows.Controls.ColumnDefinition())

        # Row 0 - Label
        self.label = Label(Content="Select views to create sheets:")
        self.label.FontFamily = FontFamily("Arial")
        self.label.FontSize = 16
        self.label.Margin = Thickness(0)
        Grid.SetRow(self.label, 0)
        grid.Children.Add(self.label)

        # Row 1 - Search Box
        self.search_box = TextBox(Height=20, FontFamily=FontFamily("Arial"), FontSize=12)
        self.search_box.TextChanged += self.search_changed
        Grid.SetRow(self.search_box, 1)
        grid.Children.Add(self.search_box)

        # Row 2 - Scrollable Checkbox Panel
        self.checkbox_panel = StackPanel(Orientation=System.Windows.Controls.Orientation.Vertical)
        scroll_viewer = ScrollViewer(Content=self.checkbox_panel, VerticalScrollBarVisibility=System.Windows.Controls.ScrollBarVisibility.Auto)
        scroll_viewer.Margin = Thickness(0, 1, 0, 1)
        Grid.SetRow(scroll_viewer, 2)
        grid.Children.Add(scroll_viewer)

        self.update_checkboxes(self.view_list)

        # Row 3 - Button Panel
        button_panel = StackPanel(Orientation=System.Windows.Controls.Orientation.Horizontal, HorizontalAlignment=HorizontalAlignment.Center, Margin=Thickness(0, 10, 0, 10))

        self.select_button = Button(Content="Select", FontFamily=FontFamily("Arial"), FontSize=12, Height=25, Margin=Thickness(10, 0, 10, 0), Width=50, HorizontalAlignment=HorizontalAlignment.Center)
        self.select_button.Click += self.select_clicked
        button_panel.Children.Add(self.select_button)

        self.check_all_button = Button(Content="Check All", FontFamily=FontFamily("Arial"), FontSize=12, Height=25, Margin=Thickness(10, 0, 10, 0), Width=70, HorizontalAlignment=HorizontalAlignment.Center)
        self.check_all_button.Click += self.check_all_clicked
        button_panel.Children.Add(self.check_all_button)

        Grid.SetRow(button_panel, 3)
        grid.Children.Add(button_panel)

        # Set window content
        self.Content = grid
        self.SizeChanged += self.on_resize

    def update_checkboxes(self, views):
        self.checkbox_panel.Children.Clear()
        self.checkboxes = []
        for view in views:
            display_name = "{} ({})".format(view.Name, view.ViewType)
            checkbox = CheckBox(Content=display_name)
            checkbox.Tag = view
            checkbox.Click += self.checkbox_clicked
            if view in self.selected_views:
                checkbox.IsChecked = True
            self.checkbox_panel.Children.Add(checkbox)
            self.checkboxes.append(checkbox)

    def search_changed(self, sender, args):
        search_text = self.search_box.Text.lower()
        filtered = [v for v in self.view_list if search_text in v.Name.lower() or search_text in str(v.ViewType).lower()]
        self.update_checkboxes(filtered)

    def check_all_clicked(self, sender, args):
        self.check_all_state = not self.check_all_state
        for cb in self.checkboxes:
            cb.IsChecked = self.check_all_state
        self.selected_views = [cb.Tag for cb in self.checkboxes if cb.IsChecked]

    def checkbox_clicked(self, sender, args):
        self.selected_views = [cb.Tag for cb in self.checkboxes if cb.IsChecked]

    def select_clicked(self, sender, args):
        self.selected_views = [cb.Tag for cb in self.checkboxes if cb.IsChecked]
        self.DialogResult = True
        self.Close()

    def on_resize(self, sender, args):
        pass

class TitleblockSelectionFilter(Window):
    def __init__(self, titleblocks):
        self.selected_titleblock = None
        self.titleblock_list = sorted(titleblocks, key=lambda x: x.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString())
        self.checkboxes = []
        self.check_all_state = False
        self.InitializeComponents()

    def InitializeComponents(self):
        self.Title = "Select Titleblock"
        self.Width = 400
        self.Height = 400
        self.MinWidth = self.Width
        self.MinHeight = self.Height
        self.ResizeMode = ResizeMode.CanResize
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen

        grid = Grid()
        grid.Margin = Thickness(5)
        for i in range(4):  # rows for: label, search box, scroll, buttons
            row = GridLength(1, GridUnitType.Star) if i == 2 else GridLength.Auto
            grid.RowDefinitions.Add(System.Windows.Controls.RowDefinition(Height=row))
        grid.ColumnDefinitions.Add(System.Windows.Controls.ColumnDefinition())

        # Row 0 - Label
        self.label = Label(Content="Select a titleblock:")
        self.label.FontFamily = FontFamily("Arial")
        self.label.FontSize = 16
        self.label.Margin = Thickness(0)
        Grid.SetRow(self.label, 0)
        grid.Children.Add(self.label)

        # Row 1 - Search Box
        self.search_box = TextBox(Height=20, FontFamily=FontFamily("Arial"), FontSize=12)
        self.search_box.TextChanged += self.search_changed
        Grid.SetRow(self.search_box, 1)
        grid.Children.Add(self.search_box)

        # Row 2 - Scrollable Checkbox Panel
        self.checkbox_panel = StackPanel(Orientation=System.Windows.Controls.Orientation.Vertical)
        scroll_viewer = ScrollViewer(Content=self.checkbox_panel, VerticalScrollBarVisibility=System.Windows.Controls.ScrollBarVisibility.Auto)
        scroll_viewer.Margin = Thickness(0, 1, 0, 1)
        Grid.SetRow(scroll_viewer, 2)
        grid.Children.Add(scroll_viewer)

        self.update_checkboxes(self.titleblock_list)

        # Row 3 - Button Panel
        button_panel = StackPanel(Orientation=System.Windows.Controls.Orientation.Horizontal, HorizontalAlignment=HorizontalAlignment.Center, Margin=Thickness(0, 10, 0, 10))

        self.select_button = Button(Content="Select", FontFamily=FontFamily("Arial"), FontSize=12, Height=25, Margin=Thickness(10, 0, 10, 0), Width=50, HorizontalAlignment=HorizontalAlignment.Center)
        self.select_button.Click += self.select_clicked
        button_panel.Children.Add(self.select_button)

        self.check_all_button = Button(Content="Check All", FontFamily=FontFamily("Arial"), FontSize=12, Height=25, Margin=Thickness(10, 0, 10, 0), Width=70, HorizontalAlignment=HorizontalAlignment.Center)
        self.check_all_button.Click += self.check_all_clicked
        button_panel.Children.Add(self.check_all_button)

        Grid.SetRow(button_panel, 3)
        grid.Children.Add(button_panel)

        # Set window content
        self.Content = grid
        self.SizeChanged += self.on_resize

    def update_checkboxes(self, titleblocks):
        self.checkbox_panel.Children.Clear()
        self.checkboxes = []
        for tb in titleblocks:
            family_name = tb.FamilyName
            type_name = tb.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            display_name = "{} - {}".format(family_name, type_name)
            checkbox = CheckBox(Content=display_name)
            checkbox.Tag = tb
            checkbox.Click += self.checkbox_clicked
            if self.selected_titleblock and type_name == self.selected_titleblock.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString():
                checkbox.IsChecked = True
            self.checkbox_panel.Children.Add(checkbox)
            self.checkboxes.append(checkbox)

    def search_changed(self, sender, args):
        search_text = self.search_box.Text.lower()
        filtered = [tb for tb in self.titleblock_list if search_text in tb.FamilyName.lower() or search_text in tb.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString().lower()]
        self.update_checkboxes(filtered)

    def check_all_clicked(self, sender, args):
        self.check_all_state = not self.check_all_state
        for cb in self.checkboxes:
            cb.IsChecked = self.check_all_state
        self.selected_titleblock = [cb.Tag for cb in self.checkboxes if cb.IsChecked][-1] if self.check_all_state else None

    def checkbox_clicked(self, sender, args):
        checked = [cb for cb in self.checkboxes if cb.IsChecked]
        if checked:
            self.selected_titleblock = checked[-1].Tag
            for cb in self.checkboxes:
                if cb != checked[-1]:
                    cb.IsChecked = False

    def select_clicked(self, sender, args):
        checked = [cb for cb in self.checkboxes if cb.IsChecked]
        if checked:
            self.selected_titleblock = checked[-1].Tag
        self.DialogResult = True
        self.Close()

    def on_resize(self, sender, args):
        pass

# Define sheet number and name input dialog
class SheetInputForm(Window):
    def __init__(self, multiple_views_selected):
        self.sheet_number = ""
        self.sheet_name = ""
        self.multiple_views_selected = multiple_views_selected
        self.InitializeComponents()

    def InitializeComponents(self):
        self.Title = "Sheet Information"
        self.Width = 300
        self.Height = 210
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen

        grid = Grid()
        grid.Margin = Thickness(10)
        for i in range(5):
            grid.RowDefinitions.Add(System.Windows.Controls.RowDefinition(Height=GridLength.Auto))
        grid.ColumnDefinitions.Add(System.Windows.Controls.ColumnDefinition())

        # Row 0 - Sheet Number Label
        label1 = Label(Content="Sheet Number:", FontFamily=FontFamily("Arial"), FontSize=12, Margin=Thickness(0, 2, 0, 2))
        Grid.SetRow(label1, 0)
        grid.Children.Add(label1)

        # Row 1 - Sheet Number Input
        default_sheet_number = 'Varies - Don\'t Change' if self.multiple_views_selected else 'Sheet Number'
        self.sheet_number_input = TextBox(Text=default_sheet_number, FontFamily=FontFamily("Arial"), FontSize=12, Height=20, Width=200, Margin=Thickness(0, 2, 0, 2))
        Grid.SetRow(self.sheet_number_input, 1)
        grid.Children.Add(self.sheet_number_input)

        # Row 2 - Sheet Name Label
        label2 = Label(Content="Sheet Name:", FontFamily=FontFamily("Arial"), FontSize=12, Margin=Thickness(0, 2, 0, 2))
        Grid.SetRow(label2, 2)
        grid.Children.Add(label2)

        # Row 3 - Sheet Name Input
        self.sheet_name_input = TextBox(Text="Sheet Name", FontFamily=FontFamily("Arial"), FontSize=12, Height=20, Width=200, Margin=Thickness(0, 2, 0, 2))
        Grid.SetRow(self.sheet_name_input, 3)
        grid.Children.Add(self.sheet_name_input)

        # Row 4 - Button
        button = Button(Content="OK", FontFamily=FontFamily("Arial"), FontSize=12, Margin=Thickness(0, 10, 0, 0), Height=30, Width=60)
        Grid.SetRow(button, 4)
        grid.Children.Add(button)
        button.Click += self.ok_clicked

        self.Content = grid

    def ok_clicked(self, sender, args):
        self.sheet_number = self.sheet_number_input.Text
        self.sheet_name = self.sheet_name_input.Text
        self.DialogResult = True
        self.Close()

# Get pre-selected views
selected_view_ids = uidoc.Selection.GetElementIds()
selected_views = []
valid_view_types = [DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan, DB.ViewType.Elevation, DB.ViewType.Section, DB.ViewType.ThreeD, DB.ViewType.DraftingView]
if selected_view_ids:
    for view_id in selected_view_ids:
        element = doc.GetElement(view_id)
        if isinstance(element, DB.View) and not element.IsTemplate and element.ViewType in valid_view_types:
            selected_views.append(element)

# If no views are pre-selected, show view selection dialog
if not selected_views:
    all_views = fec(doc).OfClass(DB.View).ToElements()
    valid_views = [v for v in all_views if not v.IsTemplate and v.ViewType in valid_view_types]
    if valid_views:
        view_form = ViewSelectionFilter(valid_views)
        if not view_form.ShowDialog() or not view_form.selected_views:
            TaskDialog.Show("Error", "No views selected.")
            sys.exit()
        selected_views = view_form.selected_views
    else:
        TaskDialog.Show("Error", "No valid views found in the document.")
        sys.exit()

# Get all titleblocks
titleblocks = fec(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType().ToElements()
if not titleblocks:
    TaskDialog.Show("Error", "No title blocks found in the document.")
    sys.exit()

# Show titleblock selection dialog
tb_form = TitleblockSelectionFilter(titleblocks)
if not tb_form.ShowDialog() or not tb_form.selected_titleblock:
    TaskDialog.Show("Error", "No title block selected.")
    sys.exit()
selected_titleblock = tb_form.selected_titleblock

# Show sheet input dialog
sheet_form = SheetInputForm(len(selected_views) > 1)
if not sheet_form.ShowDialog():
    TaskDialog.Show("Error", "No sheet information provided.")
    sys.exit()
snumber = sheet_form.sheet_number
sname = sheet_form.sheet_name

# Check if sheet number already exists (only if single view and user provided a custom sheet number)
if len(selected_views) == 1 and snumber != 'Varies - Don\'t Change':
    sheet_exists = any(sheet.SheetNumber == snumber for sheet in fec(doc).OfClass(ViewSheet).ToElements())
    if sheet_exists:
        TaskDialog.Show("Error", "Sheet with number {} already exists.".format(snumber))
        sys.exit()

# Check if selected views are already placed on any sheet
for view in selected_views:
    viewports = fec(doc).OfClass(Viewport).ToElements()
    for viewport in viewports:
        if viewport.ViewId == view.Id:
            TaskDialog.Show("Error", "View '{}' is already placed on another sheet.".format(view.Name))
            sys.exit()

# Group views by level and scope box for FloorPlan and CeilingPlan pairing
view_groups = defaultdict(list)
for view in selected_views:
    level_id = view.get_Parameter(BuiltInParameter.PLAN_VIEW_LEVEL).AsElementId() if view.ViewType in [DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan] else None
    scope_box_id = view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).AsElementId() if view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP) else None
    key = (level_id, scope_box_id)
    view_groups[key].append(view)

# Define a transaction and describe the transaction
t = Transaction(doc, 'Sheet From View')

# Begin new transaction
t.Start()
last_sheet = None
for i, (key, views) in enumerate(view_groups.items()):
    level_id, scope_box_id = key
    # Check if we have both FloorPlan and CeilingPlan with matching level and scope box
    floor_plan = next((v for v in views if v.ViewType == DB.ViewType.FloorPlan), None)
    ceiling_plan = next((v for v in views if v.ViewType == DB.ViewType.CeilingPlan), None)
    other_views = [v for v in views if v.ViewType not in [DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan]]

    if floor_plan and ceiling_plan and level_id and scope_box_id:  # Pair FloorPlan and CeilingPlan only if both level and scope box match
        # Use user-provided sheet number for single pair, else use CeilingPlan name
        current_sheet_number = snumber if len(view_groups) == 1 and snumber != 'Varies - Don\'t Change' else ceiling_plan.Name
        if any(sheet.SheetNumber == current_sheet_number for sheet in fec(doc).OfClass(ViewSheet).ToElements()):
            TaskDialog.Show("Error", "Sheet with number {} already exists.".format(current_sheet_number))
            t.RollBack()
            sys.exit()

        # Create new sheet
        SHEET = ViewSheet.Create(doc, selected_titleblock.Id)
        SHEET.Name = sname
        SHEET.SheetNumber = current_sheet_number
        last_sheet = SHEET

        # Place both CeilingPlan and FloorPlan at the same location
        x = SHEET.Outline.Max.Add(SHEET.Outline.Min).Divide(2.0)[0]
        y = SHEET.Outline.Max.Add(SHEET.Outline.Min).Divide(2.0)[1]
        ViewLocation = XYZ(x, y, 0.0)
        Viewport.Create(doc, SHEET.Id, ceiling_plan.Id, ViewLocation)
        Viewport.Create(doc, SHEET.Id, floor_plan.Id, ViewLocation)
    else:
        # Handle non-paired views or single FloorPlan/CeilingPlan separately
        for view in views:
            current_sheet_number = snumber if len(view_groups) == 1 and len(views) == 1 and snumber != 'Varies - Don\'t Change' else view.Name
            if any(sheet.SheetNumber == current_sheet_number for sheet in fec(doc).OfClass(ViewSheet).ToElements()):
                TaskDialog.Show("Error", "Sheet with number {} already exists.".format(current_sheet_number))
                t.RollBack()
                sys.exit()

            # Create new sheet
            SHEET = ViewSheet.Create(doc, selected_titleblock.Id)
            SHEET.Name = sname
            SHEET.SheetNumber = current_sheet_number
            x = SHEET.Outline.Max.Add(SHEET.Outline.Min).Divide(2.0)[0]
            y = SHEET.Outline.Max.Add(SHEET.Outline.Min).Divide(2.0)[1]
            ViewLocation = XYZ(x, y, 0.0)
            Viewport.Create(doc, SHEET.Id, view.Id, ViewLocation)
            last_sheet = SHEET

t.Commit()

# Set the active view to the last created sheet
if last_sheet:
    uidoc.RequestViewChange(last_sheet)