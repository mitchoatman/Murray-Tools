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
from Autodesk.Revit.DB import FilteredElementCollector, XYZ, BuiltInCategory, Transaction, ViewSheet, Viewport, BuiltInParameter, ElementId
from Autodesk.Revit.UI import TaskDialog, Selection
import sys

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
        self.Title = "Select Ceiling Plans"
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
        self.label = Label(Content="Select ceiling plans to add to sheets:")
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

        self.select_button = Button(Content="Select", FontFamily=FontFamily("Arial"), FontSize=12, Height=25, Margin=Thickness(10, 0, 10, 0))
        self.select_button.Click += self.select_clicked
        button_panel.Children.Add(self.select_button)

        self.check_all_button = Button(Content="Check All", FontFamily=FontFamily("Arial"), FontSize=12, Height=25, Margin=Thickness(10, 0, 10, 0))
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

class SheetSelectionFilter(Window):
    def __init__(self, sheets):
        self.selected_sheets = []
        self.sheet_list = sorted(sheets, key=lambda x: x.SheetNumber)
        self.checkboxes = []
        self.check_all_state = False
        self.InitializeComponents()

    def InitializeComponents(self):
        self.Title = "Select Sheets"
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
        self.label = Label(Content="Select sheets to add ceiling plans to:")
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

        self.update_checkboxes(self.sheet_list)

        # Row 3 - Button Panel
        button_panel = StackPanel(Orientation=System.Windows.Controls.Orientation.Horizontal, HorizontalAlignment=HorizontalAlignment.Center, Margin=Thickness(0, 10, 0, 10))

        self.select_button = Button(Content="Select", FontFamily=FontFamily("Arial"), FontSize=12, Height=25, Margin=Thickness(10, 0, 10, 0))
        self.select_button.Click += self.select_clicked
        button_panel.Children.Add(self.select_button)

        self.check_all_button = Button(Content="Check All", FontFamily=FontFamily("Arial"), FontSize=12, Height=25, Margin=Thickness(10, 0, 10, 0))
        self.check_all_button.Click += self.check_all_clicked
        button_panel.Children.Add(self.check_all_button)

        Grid.SetRow(button_panel, 3)
        grid.Children.Add(button_panel)

        # Set window content
        self.Content = grid
        self.SizeChanged += self.on_resize

    def update_checkboxes(self, sheets):
        self.checkbox_panel.Children.Clear()
        self.checkboxes = []
        for sheet in sheets:
            display_name = "{} - {}".format(sheet.SheetNumber, sheet.Name)
            checkbox = CheckBox(Content=display_name)
            checkbox.Tag = sheet
            checkbox.Click += self.checkbox_clicked
            if sheet in self.selected_sheets:
                checkbox.IsChecked = True
            self.checkbox_panel.Children.Add(checkbox)
            self.checkboxes.append(checkbox)

    def search_changed(self, sender, args):
        search_text = self.search_box.Text.lower()
        filtered = [s for s in self.sheet_list if search_text in s.SheetNumber.lower() or search_text in s.Name.lower()]
        self.update_checkboxes(filtered)

    def check_all_clicked(self, sender, args):
        self.check_all_state = not self.check_all_state
        for cb in self.checkboxes:
            cb.IsChecked = self.check_all_state
        self.selected_sheets = [cb.Tag for cb in self.checkboxes if cb.IsChecked]

    def checkbox_clicked(self, sender, args):
        self.selected_sheets = [cb.Tag for cb in self.checkboxes if cb.IsChecked]

    def select_clicked(self, sender, args):
        self.selected_sheets = [cb.Tag for cb in self.checkboxes if cb.IsChecked]
        self.DialogResult = True
        self.Close()

    def on_resize(self, sender, args):
        pass

# Get pre-selected views
selected_view_ids = uidoc.Selection.GetElementIds()
selected_views = []
if selected_view_ids:
    for view_id in selected_view_ids:
        element = doc.GetElement(view_id)
        if isinstance(element, DB.View) and not element.IsTemplate and element.ViewType == DB.ViewType.CeilingPlan:
            selected_views.append(element)

# If no ceiling plans are pre-selected, show view selection dialog with only ceiling plans
if not selected_views:
    all_views = fec(doc).OfClass(DB.View).ToElements()
    valid_views = [v for v in all_views if not v.IsTemplate and v.ViewType == DB.ViewType.CeilingPlan]
    if valid_views:
        view_form = ViewSelectionFilter(valid_views)
        if not view_form.ShowDialog() or not view_form.selected_views:
            TaskDialog.Show("Error", "No ceiling plans selected.")
            sys.exit()
        selected_views = view_form.selected_views
    else:
        TaskDialog.Show("Error", "No ceiling plans found in the document.")
        sys.exit()

# Check if selected views are already placed on any sheet
for view in selected_views:
    viewports = fec(doc).OfClass(Viewport).ToElements()
    for viewport in viewports:
        if viewport.ViewId == view.Id:
            TaskDialog.Show("Error", "Ceiling plan '{}' is already placed on another sheet.".format(view.Name))
            sys.exit()

# Show sheet selection dialog
all_sheets = fec(doc).OfClass(ViewSheet).ToElements()
if not all_sheets:
    TaskDialog.Show("Error", "No sheets found in the document.")
    sys.exit()
sheet_form = SheetSelectionFilter(all_sheets)
if not sheet_form.ShowDialog() or not sheet_form.selected_sheets:
    TaskDialog.Show("Error", "No sheets selected.")
    sys.exit()
selected_sheets = sheet_form.selected_sheets

# Check if all selected views have a scope box assigned
for view in selected_views:
    scope_box_id = view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).AsElementId() if view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP) else None
    if scope_box_id is None:
        TaskDialog.Show("Error", "No scope box assigned to ceiling plan '{}'. Please assign a scope box to all selected views.".format(view.Name))
        sys.exit()

# Define a transaction and describe the transaction
t = Transaction(doc, 'Add Ceiling Plans to Sheets')

# Begin new transaction
t.Start()
last_sheet = None
for view in selected_views:
    # Get the scope box of the ceiling plan
    scope_box_id = view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).AsElementId()
    scope_box = doc.GetElement(scope_box_id)
    scope_name = scope_box.Name if scope_box else "None"

    # Find a selected sheet with a FloorPlan view that matches the scope box
    matching_sheet = None
    for sheet in selected_sheets:
        viewports = [doc.GetElement(vp) for vp in sheet.GetAllViewports()]
        for vp in viewports:
            existing_view = doc.GetElement(vp.ViewId)
            if existing_view.ViewType != DB.ViewType.FloorPlan:
                continue  # Only match against FloorPlan views
            existing_scope_box_id = existing_view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP).AsElementId() if existing_view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP) else None
            existing_scope_name = doc.GetElement(existing_scope_box_id).Name if existing_scope_box_id else "None"
            if existing_scope_box_id is not None and existing_scope_box_id.IntegerValue == scope_box_id.IntegerValue:
                matching_sheet = sheet
                break
        if matching_sheet:
            break

    if matching_sheet:
        # Add the ceiling plan to the matching sheet at the center
        x = matching_sheet.Outline.Max.Add(matching_sheet.Outline.Min).Divide(2.0)[0]
        y = matching_sheet.Outline.Max.Add(matching_sheet.Outline.Min).Divide(2.0)[1]
        ViewLocation = XYZ(x, y, 0.0)
        Viewport.Create(doc, matching_sheet.Id, view.Id, ViewLocation)
        last_sheet = matching_sheet
    else:
        TaskDialog.Show("Warning", "No matching sheet found for ceiling plan '{}' (Scope Box: {}). Skipping.".format(
            view.Name, scope_name))

t.Commit()

# Set the active view to the last matching sheet if any
if last_sheet:
    uidoc.RequestViewChange(last_sheet)