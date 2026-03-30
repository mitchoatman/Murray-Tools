# coding: utf8
from Autodesk.Revit.DB import (
    ViewSheet,
    View,
    ViewType,
    FilteredElementCollector,
    Transaction
)
from Autodesk.Revit.UI import TaskDialog
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xaml')
import System
from System.Windows.Controls import Label, TextBox, Button, ScrollViewer, StackPanel, Grid, Orientation, CheckBox, ComboBox, RowDefinition, TextBlock
from System.Windows import Window, Thickness, SizeToContent, ResizeMode, HorizontalAlignment, GridLength, GridUnitType
from System.Windows.Media import FontFamily
from System.Windows.Controls import ScrollBarVisibility

# Define the active Revit application and document
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

class ViewRenameForm(Window):
    def __init__(self, sheets_and_views):
        self.selected_sheets_views = []
        self.sheets_views_list = sheets_and_views  # Keep full list
        self.checkboxes = []
        self.check_all_state = False
        self.prefix = ""
        self.find = ""
        self.replace = ""
        self.suffix = ""
        self.last_clicked_index = None
        self.InitializeComponents()

    def InitializeComponents(self):
        self.Title = "Rename Sheets and Views"
        self.Width = 400
        self.Height = 650
        self.MinWidth = 400
        self.MinHeight = 650
        self.ResizeMode = ResizeMode.CanResize
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.SizeToContent = SizeToContent.Manual

        grid = Grid()
        grid.Margin = Thickness(5)

        # Row definitions: label, combo box, search label, search box, scroll panel, input fields, buttons
# Correct RowDefinition/ColumnDefinition setup
        for i in range(7):
            row_def = System.Windows.Controls.RowDefinition()  # create RowDefinition instance
            if i == 4:  # the scrollable panel should expand
                row_def.Height = GridLength(1, GridUnitType.Star)
            else:
                row_def.Height = GridLength.Auto
            grid.RowDefinitions.Add(row_def)

        col_def = System.Windows.Controls.ColumnDefinition()  # create ColumnDefinition instance
        grid.ColumnDefinitions.Add(col_def)

        # Row 0 - Label
        self.label = Label(Content="Select sheets and views to rename:", FontFamily=FontFamily("Arial"), FontSize=16, Margin=Thickness(0,0,0,5))
        Grid.SetRow(self.label, 0)
        grid.Children.Add(self.label)

        # Row 1 - ComboBox filter
        self.filter_combo = ComboBox()
        self.filter_combo.Items.Add("All")
        self.filter_combo.Items.Add("Sheet Name")
        self.filter_combo.Items.Add("Sheet Number")
        self.filter_combo.Items.Add("View")
        self.filter_combo.SelectedIndex = 0
        self.filter_combo.SelectionChanged += self.filter_changed
        Grid.SetRow(self.filter_combo, 1)
        grid.Children.Add(self.filter_combo)

        # Row 2 - Search label
        search_label = Label(Content="Search:", FontFamily=FontFamily("Arial"), FontSize=12, Margin=Thickness(0,5,0,2))
        Grid.SetRow(search_label, 2)
        grid.Children.Add(search_label)

        # Row 3 - Search box
        self.search_box = TextBox(Height=20, FontFamily=FontFamily("Arial"), FontSize=12)
        self.search_box.TextChanged += self.search_changed
        self.search_box.Margin = Thickness(0,0,0,5)
        Grid.SetRow(self.search_box, 3)
        grid.Children.Add(self.search_box)

        # Row 4 - Scrollable checkbox panel
        self.checkbox_panel = StackPanel(Orientation=Orientation.Vertical)
        scroll_viewer = ScrollViewer(Content=self.checkbox_panel, VerticalScrollBarVisibility=ScrollBarVisibility.Auto)
        scroll_viewer.Margin = Thickness(0,0,0,5)
        Grid.SetRow(scroll_viewer, 4)
        grid.Children.Add(scroll_viewer)

        # Row 5 - Input fields (prefix/find/replace/suffix)
        input_panel = StackPanel(Orientation=Orientation.Vertical, Margin=Thickness(0,0,0,5))
        for lbl, attr in [("Prefix:", "prefix_input"), ("Find (Case Sensitive):", "find_input"), 
                          ("Replace:", "replace_input"), ("Suffix:", "suffix_input")]:
            label = Label(Content=lbl, FontFamily=FontFamily("Arial"), FontSize=12, Margin=Thickness(0,0,0,2))
            tb = TextBox(Height=20, FontFamily=FontFamily("Arial"), FontSize=12, Margin=Thickness(0,0,0,2))
            input_panel.Children.Add(label)
            input_panel.Children.Add(tb)
            setattr(self, attr, tb)
        Grid.SetRow(input_panel, 5)
        grid.Children.Add(input_panel)

        # Row 6 - Buttons
        button_panel = StackPanel(Orientation=Orientation.Horizontal, HorizontalAlignment=HorizontalAlignment.Center, Margin=Thickness(0,5,0,5))
        self.select_button = Button(Content="Rename", FontFamily=FontFamily("Arial"), FontSize=12, Height=25, Width=50, Margin=Thickness(10,0,10,0))
        self.select_button.Click += self.select_clicked
        self.check_all_button = Button(Content="Check All", FontFamily=FontFamily("Arial"), FontSize=12, Height=25, Width=70, Margin=Thickness(10,0,10,0))
        self.check_all_button.Click += self.check_all_clicked
        button_panel.Children.Add(self.select_button)
        button_panel.Children.Add(self.check_all_button)
        Grid.SetRow(button_panel, 6)
        grid.Children.Add(button_panel)

        self.Content = grid
        self.update_checkboxes(self.sheets_views_list)
        self.SizeChanged += self.on_resize

    def filter_changed(self, sender, args):
        self.apply_filter()

    def search_changed(self, sender, args):
        self.apply_filter()

    def apply_filter(self):
        search_text = self.search_box.Text.lower().strip()
        filter_type = self.filter_combo.SelectedItem.ToString()
        filtered = []

        for element in self.sheets_views_list:

            if isinstance(element, ViewSheet):
                name = element.Name.lower()
                number = element.SheetNumber.lower()

                # ALL = search everything
                if filter_type == "All":
                    if not search_text or search_text in name or search_text in number:
                        filtered.append(element)

                elif filter_type == "Sheet Name":
                    if not search_text or search_text in name:
                        filtered.append((element, "name"))

                elif filter_type == "Sheet Number":
                    if not search_text or search_text in number:
                        filtered.append((element, "number"))

            elif isinstance(element, View):
                display_name = "{} ({})".format(element.Name, element.ViewType).lower()

                if filter_type in ("All", "View"):
                    if not search_text or search_text in display_name:
                        filtered.append(element)

        self.update_checkboxes(filtered)

    def update_checkboxes(self, items):
        self.checkbox_panel.Children.Clear()
        self.checkboxes = []

        for item in items:
            # Sheet tuples from search/filter
            if isinstance(item, tuple):
                element, typ = item
                if typ == "name":
                    tb = TextBlock()
                    tb.Text = "Sheet Name: {}".format(element.Name)
                    cb = CheckBox(Content=tb, Tag=(element, "name"))
                else:
                    tb = TextBlock()
                    tb.Text = "Sheet Number: {}".format(element.SheetNumber)
                    cb = CheckBox(Content=tb, Tag=(element, "number"))
                cb.Click += self.checkbox_clicked
                self.checkbox_panel.Children.Add(cb)
                self.checkboxes.append(cb)

            elif isinstance(item, ViewSheet):
                # Normal view sheet, show both
                tb_name = TextBlock()
                tb_name.Text = "Sheet Name: {}".format(item.Name)
                cb_name = CheckBox(Content=tb_name, Tag=(item, "name"))
                cb_name.Click += self.checkbox_clicked

                tb_number = TextBlock()
                tb_number.Text = "Sheet Number: {}".format(item.SheetNumber)
                cb_number = CheckBox(Content=tb_number, Tag=(item, "number"))
                cb_number.Click += self.checkbox_clicked

                self.checkbox_panel.Children.Add(cb_name)
                self.checkbox_panel.Children.Add(cb_number)
                self.checkboxes.extend([cb_name, cb_number])

            elif isinstance(item, View):
                tb = TextBlock()
                tb.Text = "View: {} ({})".format(item.Name, item.ViewType)
                cb = CheckBox(Content=tb, Tag=item)
                cb.Click += self.checkbox_clicked
                self.checkbox_panel.Children.Add(cb)
                self.checkboxes.append(cb)

    def check_all_clicked(self, sender, args):
        self.check_all_state = not self.check_all_state
        for cb in self.checkboxes:
            cb.IsChecked = self.check_all_state
        self.selected_sheets_views = [cb.Tag for cb in self.checkboxes if cb.IsChecked]

    def checkbox_clicked(self, sender, args):
        try:
            # Get index of clicked TextBlock
            current_index = self.checkboxes.index(sender)

            # Check if SHIFT is held
            shift_pressed = System.Windows.Input.Keyboard.IsKeyDown(System.Windows.Input.Key.LeftShift) or \
                            System.Windows.Input.Keyboard.IsKeyDown(System.Windows.Input.Key.RightShift)

            if shift_pressed and self.last_clicked_index is not None:
                start = min(self.last_clicked_index, current_index)
                end = max(self.last_clicked_index, current_index)

                # Use the state of the clicked TextBlock
                new_state = sender.IsChecked

                for i in range(start, end + 1):
                    self.checkboxes[i].IsChecked = new_state

            # Update last clicked index
            self.last_clicked_index = current_index

            # Update selected list
            self.selected_sheets_views = [cb.Tag for cb in self.checkboxes if cb.IsChecked]

        except:
            pass

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
            for tag in self.selected_sheets_views:
                if isinstance(tag, tuple):
                    element, typ = tag
                    if typ == "name":
                        element.Name = self.prefix + element.Name.replace(self.find, self.replace) + self.suffix
                        renamed_count += 1
                    else:
                        element.SheetNumber = self.prefix + element.SheetNumber.replace(self.find, self.replace) + self.suffix
                        renamed_count += 1
                else:
                    element = tag
                    element.Name = self.prefix + element.Name.replace(self.find, self.replace) + self.suffix
                    renamed_count += 1
            t.Commit()
            TaskDialog.Show("Rename Operation", "{} elements renamed.".format(renamed_count) if renamed_count else "No sheets or views were renamed.")
        except Exception as e:
            t.RollBack()
            TaskDialog.Show("Rename Error", "Operation aborted due to error: {}".format(str(e)))
        self.Close()

    def on_resize(self, sender, args):
        pass


# Get all sheets and views
all_sheets = list(FilteredElementCollector(doc).OfClass(ViewSheet).ToElements())
all_views = [v for v in FilteredElementCollector(doc).OfClass(View).ToElements()
             if not v.IsTemplate and v.ViewType in [ViewType.FloorPlan, ViewType.CeilingPlan, ViewType.Elevation, ViewType.Section, ViewType.ThreeD, ViewType.DraftingView]]

# Combine sheets and views
sheets_and_views = sorted(all_sheets, key=lambda x: x.Name) + sorted(all_views, key=lambda x: x.Name)

if sheets_and_views:
    rename_form = ViewRenameForm(sheets_and_views)
    rename_form.ShowDialog()
else:
    TaskDialog.Show("No Elements Found", "No valid sheets or views found in the document.")