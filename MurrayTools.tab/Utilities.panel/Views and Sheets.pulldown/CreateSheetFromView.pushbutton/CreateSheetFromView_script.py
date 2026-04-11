# coding: utf8
import Autodesk
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xaml')

import System
from System.Windows.Controls import Label, TextBox, Button, ScrollViewer, StackPanel, Grid, Orientation, CheckBox, TextBlock
from System.Windows import Window, Thickness, ResizeMode, HorizontalAlignment, GridLength, GridUnitType
from System.Windows.Media import Brushes, FontFamily
from Autodesk.Revit.DB import FilteredElementCollector, XYZ, BuiltInCategory, Transaction, ViewSheet, Viewport, BuiltInParameter
from Autodesk.Revit.UI import TaskDialog
import sys
from collections import OrderedDict

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

        self.Content = grid
        self.SizeChanged += self.on_resize

    def update_checkboxes(self, views):
        self.checkbox_panel.Children.Clear()
        self.checkboxes = []
        for view in views:
            display_name = "{} ({})".format(view.Name, view.ViewType)
            tb = TextBlock()
            tb.Text = display_name
            checkbox = CheckBox(Content=tb)
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
        self.titleblock_list = sorted(
            titleblocks,
            key=lambda x: x.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        )
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
        for i in range(4):
            row = GridLength(1, GridUnitType.Star) if i == 2 else GridLength.Auto
            grid.RowDefinitions.Add(System.Windows.Controls.RowDefinition(Height=row))
        grid.ColumnDefinitions.Add(System.Windows.Controls.ColumnDefinition())

        self.label = Label(Content="Select a titleblock:")
        self.label.FontFamily = FontFamily("Arial")
        self.label.FontSize = 16
        self.label.Margin = Thickness(0)
        Grid.SetRow(self.label, 0)
        grid.Children.Add(self.label)

        self.search_box = TextBox(Height=20, FontFamily=FontFamily("Arial"), FontSize=12)
        self.search_box.TextChanged += self.search_changed
        Grid.SetRow(self.search_box, 1)
        grid.Children.Add(self.search_box)

        self.checkbox_panel = StackPanel(Orientation=System.Windows.Controls.Orientation.Vertical)
        scroll_viewer = ScrollViewer(Content=self.checkbox_panel, VerticalScrollBarVisibility=System.Windows.Controls.ScrollBarVisibility.Auto)
        scroll_viewer.Margin = Thickness(0, 1, 0, 1)
        Grid.SetRow(scroll_viewer, 2)
        grid.Children.Add(scroll_viewer)

        self.update_checkboxes(self.titleblock_list)

        button_panel = StackPanel(Orientation=System.Windows.Controls.Orientation.Horizontal, HorizontalAlignment=HorizontalAlignment.Center, Margin=Thickness(0, 10, 0, 10))

        self.select_button = Button(Content="Select", FontFamily=FontFamily("Arial"), FontSize=12, Height=25, Margin=Thickness(10, 0, 10, 0), Width=50, HorizontalAlignment=HorizontalAlignment.Center)
        self.select_button.Click += self.select_clicked
        button_panel.Children.Add(self.select_button)

        self.check_all_button = Button(Content="Check All", FontFamily=FontFamily("Arial"), FontSize=12, Height=25, Margin=Thickness(10, 0, 10, 0), Width=70, HorizontalAlignment=HorizontalAlignment.Center)
        self.check_all_button.Click += self.check_all_clicked
        button_panel.Children.Add(self.check_all_button)

        Grid.SetRow(button_panel, 3)
        grid.Children.Add(button_panel)

        self.Content = grid
        self.SizeChanged += self.on_resize

    def update_checkboxes(self, titleblocks):
        self.checkbox_panel.Children.Clear()
        self.checkboxes = []
        for tb in titleblocks:
            family_name = tb.FamilyName
            type_name = tb.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            display_name = "{} - {}".format(family_name, type_name)

            tb_display = TextBlock()
            tb_display.Text = display_name

            checkbox = CheckBox(Content=tb_display)
            checkbox.Tag = tb
            checkbox.Click += self.checkbox_clicked

            if self.selected_titleblock and tb.Id == self.selected_titleblock.Id:
                checkbox.IsChecked = True

            self.checkbox_panel.Children.Add(checkbox)
            self.checkboxes.append(checkbox)

    def search_changed(self, sender, args):
        search_text = self.search_box.Text.lower()
        filtered = []
        for tb in self.titleblock_list:
            family_name = tb.FamilyName or ""
            type_name = tb.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() or ""
            if search_text in family_name.lower() or search_text in type_name.lower():
                filtered.append(tb)
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


class SheetInputForm(Window):
    def __init__(self, default_number="", default_name=""):
        self.sheet_number = ""
        self.sheet_name = ""
        self.default_number = default_number
        self.default_name = default_name
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

        label1 = Label(Content="Sheet Number:", FontFamily=FontFamily("Arial"), FontSize=12, Margin=Thickness(0, 2, 0, 2))
        Grid.SetRow(label1, 0)
        grid.Children.Add(label1)

        self.sheet_number_input = TextBox(
            Text=self.default_number,
            FontFamily=FontFamily("Arial"),
            FontSize=12,
            Height=20,
            Width=240,
            Margin=Thickness(0, 2, 0, 2)
        )
        Grid.SetRow(self.sheet_number_input, 1)
        grid.Children.Add(self.sheet_number_input)

        label2 = Label(Content="Sheet Name:", FontFamily=FontFamily("Arial"), FontSize=12, Margin=Thickness(0, 2, 0, 2))
        Grid.SetRow(label2, 2)
        grid.Children.Add(label2)

        self.sheet_name_input = TextBox(
            Text=self.default_name,
            FontFamily=FontFamily("Arial"),
            FontSize=12,
            Height=20,
            Width=240,
            Margin=Thickness(0, 2, 0, 2)
        )
        Grid.SetRow(self.sheet_name_input, 3)
        grid.Children.Add(self.sheet_name_input)

        button = Button(Content="OK", FontFamily=FontFamily("Arial"), FontSize=12, Margin=Thickness(0, 10, 0, 0), Height=30, Width=60)
        Grid.SetRow(button, 4)
        grid.Children.Add(button)
        button.Click += self.ok_clicked

        self.Content = grid

    def ok_clicked(self, sender, args):
        self.sheet_number = self.sheet_number_input.Text.strip()
        self.sheet_name = self.sheet_name_input.Text.strip()

        if not self.sheet_number or not self.sheet_name:
            TaskDialog.Show("Error", "Sheet number and sheet name are required.")
            return

        self.DialogResult = True
        self.Close()


class MultiSheetInputForm(Window):
    def __init__(self, sheet_targets):
        self.sheet_targets = sheet_targets
        self.sheet_data = []
        self.row_inputs = []
        self.InitializeComponents()

    def InitializeComponents(self):
        self.Title = "Sheet Information"
        self.Width = 600
        self.Height = 400
        self.MinWidth = 500
        self.MinHeight = 400
        self.ResizeMode = ResizeMode.CanResize
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen

        root = Grid()
        root.Margin = Thickness(10)
        root.RowDefinitions.Add(System.Windows.Controls.RowDefinition(Height=GridLength.Auto))
        root.RowDefinitions.Add(System.Windows.Controls.RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        root.RowDefinitions.Add(System.Windows.Controls.RowDefinition(Height=GridLength.Auto))

        header = Grid()
        header.Margin = Thickness(0, 0, 18, 6)
        header.ColumnDefinitions.Add(System.Windows.Controls.ColumnDefinition(Width=GridLength(180)))
        header.ColumnDefinitions.Add(System.Windows.Controls.ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))

        num_header = TextBlock(Text="NUMBER", FontFamily=FontFamily("Arial"), FontSize=18, Margin=Thickness(4, 0, 10, 0))
        name_header = TextBlock(Text="NAME", FontFamily=FontFamily("Arial"), FontSize=18, Margin=Thickness(4, 0, 0, 0))

        Grid.SetColumn(num_header, 0)
        Grid.SetColumn(name_header, 1)

        header.Children.Add(num_header)
        header.Children.Add(name_header)

        Grid.SetRow(header, 0)
        root.Children.Add(header)

        self.rows_panel = StackPanel(Orientation=Orientation.Vertical)

        for item in self.sheet_targets:
            row = Grid()
            row.Margin = Thickness(0, 0, 0, 4)
            row.ColumnDefinitions.Add(System.Windows.Controls.ColumnDefinition(Width=GridLength(180)))
            row.ColumnDefinitions.Add(System.Windows.Controls.ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))

            num_box = TextBox(
                Text=item["default_number"],
                Height=16,
                Padding=Thickness(2, 0, 2, 0),
                Margin=Thickness(0, 0, 12, 0),
                FontFamily=FontFamily("Arial"),
                FontSize=11
            )

            name_box = TextBox(
                Text=item["default_name"],
                Height=16,
                Padding=Thickness(2, 0, 2, 0),
                FontFamily=FontFamily("Arial"),
                FontSize=12
            )

            Grid.SetColumn(num_box, 0)
            Grid.SetColumn(name_box, 1)

            row.Children.Add(num_box)
            row.Children.Add(name_box)
            self.rows_panel.Children.Add(row)

            self.row_inputs.append({
                "item": item,
                "number_box": num_box,
                "name_box": name_box
            })

        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
        scroll.Content = self.rows_panel

        Grid.SetRow(scroll, 1)
        root.Children.Add(scroll)

        button_panel = StackPanel(
            Orientation=Orientation.Horizontal,
            HorizontalAlignment=HorizontalAlignment.Right,
            Margin=Thickness(0, 10, 0, 0)
        )

        ok_button = Button(Content="OK", Width=80, Height=30, Margin=Thickness(6, 0, 0, 0))
        ok_button.Click += self.ok_clicked
        button_panel.Children.Add(ok_button)

        cancel_button = Button(Content="Cancel", Width=80, Height=30, Margin=Thickness(6, 0, 0, 0))
        cancel_button.Click += self.cancel_clicked
        button_panel.Children.Add(cancel_button)

        Grid.SetRow(button_panel, 2)
        root.Children.Add(button_panel)

        self.Content = root

    def ok_clicked(self, sender, args):
        self.sheet_data = []
        numbers = []

        for row in self.row_inputs:
            sheet_number = row["number_box"].Text.strip()
            sheet_name = row["name_box"].Text.strip()

            if not sheet_number or not sheet_name:
                TaskDialog.Show("Error", "All sheet number and sheet name fields must be filled in.")
                return

            numbers.append(sheet_number)

            self.sheet_data.append({
                "label": row["item"]["label"],
                "views": row["item"]["views"],
                "sheet_number": sheet_number,
                "sheet_name": sheet_name
            })

        duplicates = []
        for n in set(numbers):
            if numbers.count(n) > 1:
                duplicates.append(n)

        if duplicates:
            TaskDialog.Show("Error", "Duplicate sheet numbers entered:\n{}".format("\n".join(sorted(duplicates))))
            self.sheet_data = []
            return

        self.DialogResult = True
        self.Close()

    def cancel_clicked(self, sender, args):
        self.DialogResult = False
        self.Close()

def build_sheet_targets(selected_views):
    grouped = OrderedDict()

    for view in selected_views:
        level_id = -1
        scope_id = -1

        if view.ViewType in [DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan]:
            level_param = view.get_Parameter(BuiltInParameter.PLAN_VIEW_LEVEL)
            if level_param:
                level_id = level_param.AsElementId().IntegerValue

        scope_param = view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)
        if scope_param:
            scope_id = scope_param.AsElementId().IntegerValue

        key = (level_id, scope_id)
        if key not in grouped:
            grouped[key] = []
        grouped[key].append(view)

    sheet_targets = []

    for key, views in grouped.items():
        level_id, scope_id = key

        floor_plan = next((v for v in views if v.ViewType == DB.ViewType.FloorPlan), None)
        ceiling_plan = next((v for v in views if v.ViewType == DB.ViewType.CeilingPlan), None)

        if floor_plan and ceiling_plan and level_id != -1 and scope_id != -1:
            sheet_targets.append({
                "label": "{} + {}".format(ceiling_plan.Name, floor_plan.Name),
                "views": [ceiling_plan, floor_plan],
                "default_number": ceiling_plan.Name,
                "default_name": ceiling_plan.Name
            })

            for view in views:
                if view.Id != floor_plan.Id and view.Id != ceiling_plan.Id:
                    sheet_targets.append({
                        "label": view.Name,
                        "views": [view],
                        "default_number": view.Name,
                        "default_name": view.Name
                    })
        else:
            for view in views:
                sheet_targets.append({
                    "label": view.Name,
                    "views": [view],
                    "default_number": view.Name,
                    "default_name": view.Name
                })

    return sheet_targets


# Get pre-selected views
selected_view_ids = uidoc.Selection.GetElementIds()
selected_views = []

valid_view_types = [
    DB.ViewType.FloorPlan,
    DB.ViewType.CeilingPlan,
    DB.ViewType.Elevation,
    DB.ViewType.Section,
    DB.ViewType.ThreeD,
    DB.ViewType.DraftingView
]

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

# Check if selected views are already placed on any sheet
all_viewports = list(fec(doc).OfClass(Viewport).ToElements())
views_on_sheets = []
valid_views = []

for view in selected_views:
    is_on_sheet = False
    for viewport in all_viewports:
        if viewport.ViewId == view.Id:
            views_on_sheets.append(view.Name)
            is_on_sheet = True
            break
    if not is_on_sheet:
        valid_views.append(view)

# Show warning if any views are already on sheets
if views_on_sheets:
    warning_message = "The following views are already on sheets and will be skipped:\n\n" + "\n".join(views_on_sheets)
    if valid_views:
        warning_message += "\n\nContinuing with {} remaining view(s).".format(len(valid_views))
        TaskDialog.Show("Warning", warning_message)
    else:
        warning_message += "\n\nNo valid views remaining to process."
        TaskDialog.Show("Error", warning_message)
        sys.exit()

# Exit if no valid views remain
if not valid_views:
    TaskDialog.Show("Error", "No valid views to process.")
    sys.exit()

# Update selected_views to only include valid views
selected_views = valid_views

# Build target sheet list first
sheet_targets = build_sheet_targets(selected_views)

# Show input dialog
if len(sheet_targets) == 1:
    sheet_form = SheetInputForm(
        sheet_targets[0]["default_number"],
        sheet_targets[0]["default_name"]
    )
    if not sheet_form.ShowDialog():
        # TaskDialog.Show("Error", "No sheet information provided.")
        sys.exit()

    sheet_targets[0]["sheet_number"] = sheet_form.sheet_number
    sheet_targets[0]["sheet_name"] = sheet_form.sheet_name
else:
    multi_form = MultiSheetInputForm(sheet_targets)
    if not multi_form.ShowDialog():
        # TaskDialog.Show("Error", "No sheet information provided.")
        sys.exit()

    sheet_targets = multi_form.sheet_data

# Check against existing sheet numbers
existing_sheet_numbers = set(sheet.SheetNumber for sheet in fec(doc).OfClass(ViewSheet).ToElements())
for target in sheet_targets:
    if target["sheet_number"] in existing_sheet_numbers:
        TaskDialog.Show("Error", "Sheet with number {} already exists.".format(target["sheet_number"]))
        sys.exit()

# Create sheets
t = Transaction(doc, 'Sheet From View')
t.Start()

last_sheet = None

try:
    for target in sheet_targets:
        current_sheet_number = target["sheet_number"]
        current_sheet_name = target["sheet_name"]

        if current_sheet_number in existing_sheet_numbers:
            TaskDialog.Show("Error", "Sheet with number {} already exists.".format(current_sheet_number))
            t.RollBack()
            sys.exit()

        SHEET = ViewSheet.Create(doc, selected_titleblock.Id)
        SHEET.Name = current_sheet_name
        SHEET.SheetNumber = current_sheet_number

        x = SHEET.Outline.Max.Add(SHEET.Outline.Min).Divide(2.0)[0]
        y = SHEET.Outline.Max.Add(SHEET.Outline.Min).Divide(2.0)[1]
        ViewLocation = XYZ(x, y, 0.0)

        for view in target["views"]:
            Viewport.Create(doc, SHEET.Id, view.Id, ViewLocation)

        existing_sheet_numbers.add(current_sheet_number)
        last_sheet = SHEET

    t.Commit()

except Exception as ex:
    if t.HasStarted():
        t.RollBack()
    TaskDialog.Show("Error", str(ex))
    sys.exit()

# Set the active view to the last created sheet
if last_sheet:
    uidoc.RequestViewChange(last_sheet)