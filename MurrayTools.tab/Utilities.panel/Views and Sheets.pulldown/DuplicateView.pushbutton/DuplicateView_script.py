# coding: utf8
import Autodesk
from Autodesk.Revit.DB import Transaction, ViewDuplicateOption, FilteredElementCollector, ElementId, View
from Autodesk.Revit.UI import Selection, TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
import sys
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xaml')
import System
from System.Windows.Controls import Label, TextBox, Button, ComboBox, Grid, StackPanel, Orientation
from System.Windows import Window, Thickness, HorizontalAlignment, VerticalAlignment, GridLength, GridUnitType
from System.Windows.Media import Brushes, FontFamily
from System.Windows import SizeToContent, ResizeMode, WindowStartupLocation
from collections import namedtuple

#---Define the active Revit application and document
DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

#---Custom dialog for view selection
CategoryOption = namedtuple('CategoryOption', ['name', 'revit_view'])

class FilterSelectionByCategory(Window):
    def __init__(self, item_list, title, label_text):
        self.selected_items = []
        self.item_list = sorted(item_list, key=lambda x: x.Name)
        self.checkboxes = []
        self.check_all_state = False
        self.title = title
        self.label_text = label_text
        self.InitializeComponents()

    def InitializeComponents(self):
        self.Title = self.title
        self.Width = 400
        self.Height = 400
        self.MinWidth = self.Width
        self.MinHeight = self.Height
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen

        grid = Grid()
        grid.Margin = Thickness(5)
        for i in range(4):  # rows for: label, search box, scroll, buttons
            row = GridLength(1, GridUnitType.Star) if i == 2 else GridLength.Auto
            grid.RowDefinitions.Add(System.Windows.Controls.RowDefinition(Height=row))
        grid.ColumnDefinitions.Add(System.Windows.Controls.ColumnDefinition())

        # Row 0 - Label
        self.label = Label(Content=self.label_text)
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
        scroll_viewer = System.Windows.Controls.ScrollViewer(Content=self.checkbox_panel, VerticalScrollBarVisibility=System.Windows.Controls.ScrollBarVisibility.Auto)
        scroll_viewer.Margin = Thickness(0, 1, 0, 1)
        Grid.SetRow(scroll_viewer, 2)
        grid.Children.Add(scroll_viewer)

        self.update_checkboxes(self.item_list)

        # Row 3 - Button Panel
        button_panel = StackPanel(Orientation=System.Windows.Controls.Orientation.Horizontal, HorizontalAlignment=HorizontalAlignment.Center, Margin=Thickness(0, 10, 0, 10))

        self.select_button = Button(Content="Select", FontFamily=FontFamily("Arial"), FontSize=12, Height=25, Margin=Thickness(10, 0, 10, 0), Width=50)
        self.select_button.Click += self.select_clicked
        button_panel.Children.Add(self.select_button)

        self.check_all_button = Button(Content="Check All", FontFamily=FontFamily("Arial"), FontSize=12, Height=25, Margin=Thickness(10, 0, 10, 0), Width=70)
        self.check_all_button.Click += self.check_all_clicked
        button_panel.Children.Add(self.check_all_button)

        Grid.SetRow(button_panel, 3)
        grid.Children.Add(button_panel)

        # Set window content
        self.Content = grid
        self.SizeChanged += self.on_resize

    def update_checkboxes(self, items):
        self.checkbox_panel.Children.Clear()
        self.checkboxes = []
        for item in items:
            # Display view name with view type
            display_name = "{} ({})".format(item.Name, item.ViewType)
            checkbox = System.Windows.Controls.CheckBox(Content=display_name)
            checkbox.Tag = item
            checkbox.Click += self.checkbox_clicked
            if display_name in self.selected_items:
                checkbox.IsChecked = True
            self.checkbox_panel.Children.Add(checkbox)
            self.checkboxes.append(checkbox)

    def search_changed(self, sender, args):
        search_text = self.search_box.Text.lower()
        filtered = [c for c in self.item_list if search_text in c.Name.lower()]
        self.update_checkboxes(filtered)

    def check_all_clicked(self, sender, args):
        self.check_all_state = not self.check_all_state
        for cb in self.checkboxes:
            cb.IsChecked = self.check_all_state
        self.selected_items = [cb.Content for cb in self.checkboxes if cb.IsChecked]

    def checkbox_clicked(self, sender, args):
        self.selected_items = [cb.Content for cb in self.checkboxes if cb.IsChecked]

    def select_clicked(self, sender, args):
        self.selected_items = [cb.Content for cb in self.checkboxes if cb.IsChecked]
        self.DialogResult = True
        self.Close()

    def on_resize(self, sender, args):
        pass

#---Custom dialog for duplication options
class DuplicateOptionsDialog(Window):
    def __init__(self):
        self.DuplicateMode = None
        self.NumberOfCopies = None
        self.InitializeComponents()

    def InitializeComponents(self):
        self.Title = "Duplicate Views"
        self.Width = 300
        self.Height = 200
        self.MinWidth = self.Width
        self.MinHeight = self.Height
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen

        grid = Grid()
        grid.Margin = Thickness(10)
        for i in range(5):  # rows for: label, combobox, label, textbox, button
            grid.RowDefinitions.Add(System.Windows.Controls.RowDefinition(Height=GridLength.Auto))
        grid.ColumnDefinitions.Add(System.Windows.Controls.ColumnDefinition())

        # Row 0 - Copy Method Label
        self.label_copy_method = Label(Content="Copy Method", FontFamily=FontFamily("Arial"), FontSize=12)
        Grid.SetRow(self.label_copy_method, 0)
        grid.Children.Add(self.label_copy_method)

        # Row 1 - Copy Method ComboBox
        self.combo_box = ComboBox(FontFamily=FontFamily("Arial"), FontSize=12)
        self.combo_box.Items.Add("Duplicate")
        self.combo_box.Items.Add("Duplicate with Detailing")
        self.combo_box.Items.Add("Duplicate as a Dependent")
        self.combo_box.SelectedIndex = 0  # Default to "Duplicate"
        Grid.SetRow(self.combo_box, 1)
        grid.Children.Add(self.combo_box)

        # Row 2 - Number of Copies Label
        self.label_num_copies = Label(Content="Number of Copies", FontFamily=FontFamily("Arial"), FontSize=12, Margin=Thickness(0, 10, 0, 0))
        Grid.SetRow(self.label_num_copies, 2)
        grid.Children.Add(self.label_num_copies)

        # Row 3 - Number of Copies TextBox
        self.text_box = TextBox(Text="1", FontFamily=FontFamily("Arial"), FontSize=12, Height=25)
        Grid.SetRow(self.text_box, 3)
        grid.Children.Add(self.text_box)

        # Row 4 - OK Button
        self.ok_button = Button(Content="OK", FontFamily=FontFamily("Arial"), FontSize=12, Height=25, Width=100, Margin=Thickness(10, 10, 10, 0), HorizontalAlignment=HorizontalAlignment.Center)
        self.ok_button.Click += self.ok_clicked
        Grid.SetRow(self.ok_button, 4)
        grid.Children.Add(self.ok_button)

        self.Content = grid

    def ok_clicked(self, sender, args):
        self.DuplicateMode = self.combo_box.SelectedItem
        try:
            self.NumberOfCopies = int(self.text_box.Text)
            if self.NumberOfCopies <= 0:
                TaskDialog.Show("Error", "Number of copies must be a positive integer.")
                return
            self.DialogResult = True
            self.Close()
        except ValueError:
            TaskDialog.Show("Error", "Number of copies must be a valid integer.")
            return

#---Get Selected Views
# Check for pre-selected views
selected_view_ids = uidoc.Selection.GetElementIds()
selected_views = []
if selected_view_ids:
    for view_id in selected_view_ids:
        element = doc.GetElement(view_id)
        if isinstance(element, DB.View) and element.CanBePrinted:
            selected_views.append(element)

# If no views are selected, prompt with dialog
if not selected_views:
    views = [v for v in FilteredElementCollector(doc).OfClass(DB.View).WhereElementIsNotElementType().ToElements() if v.CanBePrinted]
    if views:
        form = FilterSelectionByCategory(views, "Select Views to Duplicate", "Search and select views:")
        if form.ShowDialog():
            selected_views = [x for x in views if "{} ({})".format(x.Name, x.ViewType) in form.selected_items]
        else:
            TaskDialog.Show("Error", "No views selected. Please try again.")
            sys.exit()
    else:
        TaskDialog.Show("Error", "No printable views found in the document.")
        sys.exit()

#---Display dialog to get duplication options
options_form = DuplicateOptionsDialog()
if not options_form.ShowDialog():
    TaskDialog.Show("Error", "No duplication options selected. Please try again.")
    sys.exit()

try:
    #---Convert dialog input into variables
    DuplicateMode = options_form.DuplicateMode
    NumberOfCopies = options_form.NumberOfCopies

    # Start transaction to perform the duplication in bulk
    t = Transaction(doc, 'Duplicate View(s)')
    t.Start()

    # Dictionary to store views and duplication methods
    duplication_tasks = {}

    for view in selected_views:
        for i in range(NumberOfCopies):
            if DuplicateMode == 'Duplicate':
                if view.ViewType == 'Legend':
                    duplication_tasks[view] = ViewDuplicateOption.WithDetailing
                else:
                    duplication_tasks[view] = ViewDuplicateOption.Duplicate

            elif DuplicateMode == 'Duplicate with Detailing':
                if type(view) == 'Schedule':
                    duplication_tasks[view] = ViewDuplicateOption.Duplicate
                else:
                    duplication_tasks[view] = ViewDuplicateOption.WithDetailing

            elif DuplicateMode == 'Duplicate as a Dependent':
                duplication_tasks[view] = ViewDuplicateOption.AsDependent

    # Process all duplications in the same transaction
    for view, option in duplication_tasks.items():
        for _ in range(NumberOfCopies):
            view.Duplicate(option)

    t.Commit()
except Exception as e:
    t.RollBack()
    TaskDialog.Show("Error", "An error occurred: {}".format(str(e)))
    sys.exit()