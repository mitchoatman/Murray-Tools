# coding: utf8
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType, PickBoxStyle, Selection
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xaml')
import System
from System.Windows.Controls import Label, TextBox, Button, ScrollViewer, StackPanel, Grid, Orientation, CheckBox
from System.Windows import Window, Thickness, SizeToContent, ResizeMode, HorizontalAlignment, VerticalAlignment, GridLength, GridUnitType
from System.Windows.Media import Brushes, FontFamily

# Define the active Revit application and document
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

# Define the view selection and rename dialog
class ViewRenameForm(Window):
    def __init__(self, sheets_and_views):
        self.selected_sheets_views = []
        self.sheets_views_list = sorted(sheets_and_views, key=lambda x: x.Name)
        self.checkboxes = []
        self.check_all_state = False
        self.prefix = ""
        self.find = ""
        self.replace = ""
        self.suffix = ""
        self.InitializeComponents()

    def InitializeComponents(self):
        self.Title = "Rename Sheets and Views"
        self.Width = 400
        self.Height = 650
        self.MinWidth = self.Width
        self.MinHeight = self.Height
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.SizeToContent = SizeToContent.Manual

        grid = System.Windows.Controls.Grid()
        grid.Margin = Thickness(5)
        # Rows: label, search box, scrollable checklist, input fields panel, button panel
        for i in range(5):
            row = GridLength(1, GridUnitType.Star) if i == 2 else GridLength.Auto
            grid.RowDefinitions.Add(System.Windows.Controls.RowDefinition(Height=row))
        grid.ColumnDefinitions.Add(System.Windows.Controls.ColumnDefinition())

        # Row 0 - Label
        self.label = Label(Content="Select sheets and views to rename:")
        self.label.FontFamily = FontFamily("Arial")
        self.label.FontSize = 16
        self.label.Margin = Thickness(0, 0, 0, 5)
        System.Windows.Controls.Grid.SetRow(self.label, 0)
        grid.Children.Add(self.label)

        # Row 1 - Search Box
        self.search_box = TextBox(Height=20, FontFamily=FontFamily("Arial"), FontSize=12)
        self.search_box.TextChanged += self.search_changed
        self.search_box.Margin = Thickness(0, 0, 0, 5)
        System.Windows.Controls.Grid.SetRow(self.search_box, 1)
        grid.Children.Add(self.search_box)

        # Row 2 - Scrollable Checkbox Panel
        self.checkbox_panel = StackPanel(Orientation=System.Windows.Controls.Orientation.Vertical)
        scroll_viewer = ScrollViewer(Content=self.checkbox_panel, VerticalScrollBarVisibility=System.Windows.Controls.ScrollBarVisibility.Auto)
        scroll_viewer.Margin = Thickness(0, 0, 0, 5)
        scroll_viewer.MaxHeight = 300
        System.Windows.Controls.Grid.SetRow(scroll_viewer, 2)
        grid.Children.Add(scroll_viewer)

        self.update_checkboxes(self.sheets_views_list)

        # Row 3 - Input Fields Panel
        input_panel = StackPanel(Orientation=System.Windows.Controls.Orientation.Vertical, Margin=Thickness(0, 0, 0, 5))
        
        # Prefix
        prefix_label = Label(Content="Prefix:", FontFamily=FontFamily("Arial"), FontSize=12, Margin=Thickness(0, 0, 0, 2))
        input_panel.Children.Add(prefix_label)
        self.prefix_input = TextBox(Height=20, FontFamily=FontFamily("Arial"), FontSize=12, Margin=Thickness(0, 0, 0, 2))
        input_panel.Children.Add(self.prefix_input)

        # Find
        find_label = Label(Content="Find (Case Sensitive):", FontFamily=FontFamily("Arial"), FontSize=12, Margin=Thickness(0, 0, 0, 2))
        input_panel.Children.Add(find_label)
        self.find_input = TextBox(Height=20, FontFamily=FontFamily("Arial"), FontSize=12, Margin=Thickness(0, 0, 0, 2))
        input_panel.Children.Add(self.find_input)

        # Replace
        replace_label = Label(Content="Replace:", FontFamily=FontFamily("Arial"), FontSize=12, Margin=Thickness(0, 0, 0, 2))
        input_panel.Children.Add(replace_label)
        self.replace_input = TextBox(Height=20, FontFamily=FontFamily("Arial"), FontSize=12, Margin=Thickness(0, 0, 0, 2))
        input_panel.Children.Add(self.replace_input)

        # Suffix
        suffix_label = Label(Content="Suffix:", FontFamily=FontFamily("Arial"), FontSize=12, Margin=Thickness(0, 0, 0, 2))
        input_panel.Children.Add(suffix_label)
        self.suffix_input = TextBox(Height=20, FontFamily=FontFamily("Arial"), FontSize=12, Margin=Thickness(0, 0, 0, 2))
        input_panel.Children.Add(self.suffix_input)

        System.Windows.Controls.Grid.SetRow(input_panel, 3)
        grid.Children.Add(input_panel)

        # Row 4 - Button Panel
        button_panel = StackPanel(Orientation=System.Windows.Controls.Orientation.Horizontal, HorizontalAlignment=HorizontalAlignment.Center, Margin=Thickness(0, 5, 0, 5))
        self.select_button = Button(Content="Rename", FontFamily=FontFamily("Arial"), FontSize=12, Height=25, Margin=Thickness(10, 0, 10, 0), Width=50, HorizontalAlignment=HorizontalAlignment.Center)
        self.select_button.Click += self.select_clicked
        button_panel.Children.Add(self.select_button)

        self.check_all_button = Button(Content="Check All", FontFamily=FontFamily("Arial"), FontSize=12, Height=25, Margin=Thickness(10, 0, 10, 0), Width=70, HorizontalAlignment=HorizontalAlignment.Center)
        self.check_all_button.Click += self.check_all_clicked
        button_panel.Children.Add(self.check_all_button)

        System.Windows.Controls.Grid.SetRow(button_panel, 4)
        grid.Children.Add(button_panel)

        # Set window content
        self.Content = grid
        self.SizeChanged += self.on_resize

    def update_checkboxes(self, sheets_views):
        self.checkbox_panel.Children.Clear()
        self.checkboxes = []
        for element in sheets_views:
            if isinstance(element, ViewSheet):
                display_name = "Sheet: {}".format(element.Name)
            else:
                display_name = "View: {} ({})".format(element.Name, element.ViewType)
            checkbox = CheckBox(Content=display_name)
            checkbox.Tag = element
            checkbox.Click += self.checkbox_clicked
            if element in self.selected_sheets_views:
                checkbox.IsChecked = True
            self.checkbox_panel.Children.Add(checkbox)
            self.checkboxes.append(checkbox)

    def search_changed(self, sender, args):
        search_text = self.search_box.Text.lower()
        filtered = []
        for element in self.sheets_views_list:
            # Create display name for search
            if isinstance(element, ViewSheet):
                display_name = "Sheet: {}".format(element.Name)
            else:
                display_name = "View: {} ({})".format(element.Name, element.ViewType)
            # Check if search text matches name, display name, or ViewType
            if (search_text in element.Name.lower() or
                search_text in display_name.lower() or
                (not isinstance(element, ViewSheet) and search_text in str(element.ViewType).lower())):
                filtered.append(element)
        self.update_checkboxes(filtered)

    def check_all_clicked(self, sender, args):
        self.check_all_state = not self.check_all_state
        for cb in self.checkboxes:
            cb.IsChecked = self.check_all_state
        self.selected_sheets_views = [cb.Tag for cb in self.checkboxes if cb.IsChecked]

    def checkbox_clicked(self, sender, args):
        self.selected_sheets_views = [cb.Tag for cb in self.checkboxes if cb.IsChecked]

    def select_clicked(self, sender, args):
        self.selected_sheets_views = [cb.Tag for cb in self.checkboxes if cb.IsChecked]
        if not self.selected_sheets_views:
            TaskDialog.Show("Rename Operation", "No sheets or views selected for renaming.")
            return

        self.prefix = self.prefix_input.Text
        self.find = self.find_input.Text
        self.replace = self.replace_input.Text
        self.suffix = self.suffix_input.Text

        try:
            t = Transaction(doc, 'Rename Sheets and Views')
            t.Start()
            renamed_count = 0
            for element in self.selected_sheets_views:
                new_name = self.prefix + element.Name.replace(self.find, self.replace) + self.suffix
                for i in range(20):
                    try:
                        element.Name = new_name
                        renamed_count += 1
                        break
                    except:
                        new_name += "*"
            t.Commit()
            if renamed_count > 0:
                TaskDialog.Show("Rename Operation", "Elements renamed.".format(renamed_count))
            else:
                TaskDialog.Show("Rename Operation", "No sheets or views were renamed.")
        except Exception as e:
            TaskDialog.Show("Rename Error", "Operation aborted due to an error: {}".format(str(e)))
        self.Close()

    def on_resize(self, sender, args):
        pass

# Get all sheets and views
all_sheets = list(FilteredElementCollector(doc).OfClass(ViewSheet).ToElements())  # Convert to Python list
all_views = FilteredElementCollector(doc).OfClass(View).ToElements()
valid_view_types = [ViewType.FloorPlan, ViewType.CeilingPlan, ViewType.Elevation, ViewType.Section, ViewType.ThreeD, ViewType.DraftingView]
valid_views = [v for v in all_views if not v.IsTemplate and v.ViewType in valid_view_types]
sheets_and_views = all_sheets + valid_views

if sheets_and_views:
    # Show view rename dialog
    rename_form = ViewRenameForm(sheets_and_views)
    rename_form.ShowDialog()
else:
    TaskDialog.Show("No Elements Found", "No valid sheets or views found in the document.")