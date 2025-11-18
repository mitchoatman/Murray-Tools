# -*- coding: utf-8 -*-
from collections import namedtuple
from System.Collections.Generic import List
from Autodesk.Revit import DB
from Autodesk.Revit.DB import ViewSheet, Viewport, FilteredElementCollector, XYZ
from Autodesk.Revit.UI import TaskDialog
import clr, sys
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Xaml')
import System
from System.Windows.Controls import Label, TextBox, Button, ScrollViewer, StackPanel, Grid, Orientation
from System.Windows import Window, Thickness, SizeToContent, ResizeMode, HorizontalAlignment, GridLength, GridUnitType

doc = __revit__.ActiveUIDocument.Document

SheetOption = namedtuple('SheetOption', ['name', 'sheet'])

class SheetSelectionWindow(Window):
    def __init__(self, sheet_list, multiselect=False, title="Select Sheet"):
        self.selected_sheets = []
        self.sheet_list = sorted(sheet_list, key=lambda x: x.SheetNumber + " - " + x.Name)
        self.multiselect = multiselect
        self.checkboxes = []
        self.check_all_state = False
        self.InitializeComponents(title)

    def InitializeComponents(self, title):
        self.Title = title
        self.Width = 500
        self.Height = 400
        self.MinWidth = self.Width
        self.MinHeight = self.Height
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen

        grid = Grid()
        grid.Margin = Thickness(5)
        for i in range(4):  # rows for: label, search box, scroll, buttons
            row = GridLength(1, GridUnitType.Star) if i == 2 else GridLength.Auto
            grid.RowDefinitions.Add(System.Windows.Controls.RowDefinition(Height=row))
        grid.ColumnDefinitions.Add(System.Windows.Controls.ColumnDefinition())

        # Row 0 - Label
        self.label = Label(Content="Search and select sheets:")
        self.label.FontFamily = System.Windows.Media.FontFamily("Arial")
        self.label.FontSize = 16
        self.label.Margin = Thickness(0)
        Grid.SetRow(self.label, 0)
        grid.Children.Add(self.label)

        # Row 1 - Search Box
        self.search_box = TextBox(Height=20, FontFamily=System.Windows.Media.FontFamily("Arial"), FontSize=12)
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

        self.select_button = Button(Content="Select", FontFamily=System.Windows.Media.FontFamily("Arial"), FontSize=12, Height=25, Margin=Thickness(0, 0, 20, 0))
        self.select_button.Click += self.select_clicked
        button_panel.Children.Add(self.select_button)

        if self.multiselect:
            self.check_all_button = Button(Content="Check All", FontFamily=System.Windows.Media.FontFamily("Arial"), FontSize=12, Height=25)
            self.check_all_button.Click += self.check_all_clicked
            button_panel.Children.Add(self.check_all_button)

        Grid.SetRow(button_panel, 3)
        grid.Children.Add(button_panel)

        # Set window content
        self.Content = grid

    def update_checkboxes(self, sheets):
        self.checkbox_panel.Children.Clear()
        self.checkboxes = []
        for sheet in sheets:
            display_name = sheet.SheetNumber + " - " + sheet.Name
            checkbox = System.Windows.Controls.CheckBox(Content=display_name)
            checkbox.Tag = sheet
            checkbox.Click += self.checkbox_clicked
            if display_name in self.selected_sheets:
                checkbox.IsChecked = True
            self.checkbox_panel.Children.Add(checkbox)
            self.checkboxes.append(checkbox)

    def search_changed(self, sender, args):
        search_text = self.search_box.Text.lower()
        filtered = [s for s in self.sheet_list if search_text in (s.SheetNumber + " - " + s.Name).lower()]
        self.update_checkboxes(filtered)

    def check_all_clicked(self, sender, args):
        self.check_all_state = not self.check_all_state
        for cb in self.checkboxes:
            cb.IsChecked = self.check_all_state
        self.selected_sheets = [cb.Tag for cb in self.checkboxes if cb.IsChecked]

    def checkbox_clicked(self, sender, args):
        if not self.multiselect:
            for cb in self.checkboxes:
                cb.IsChecked = False
            sender.IsChecked = True
        self.selected_sheets = [cb.Tag for cb in self.checkboxes if cb.IsChecked]

    def select_clicked(self, sender, args):
        self.selected_sheets = [cb.Tag for cb in self.checkboxes if cb.IsChecked]
        self.DialogResult = True
        self.Close()

# Get all sheets
all_sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()

# Step 1 - User selects template sheet
form = SheetSelectionWindow(all_sheets, multiselect=False, title="Select Template Sheet (To Get View Position)")
if not form.ShowDialog() or not form.selected_sheets:
    sys.exit()
template_sheet = form.selected_sheets[0]

# Step 2 - Get first viewport's position from template sheet
template_vp = None
for vp_id in template_sheet.GetAllViewports():
    template_vp = doc.GetElement(vp_id)
    break

if not template_vp:
    TaskDialog.Show("Error", "No views found on the selected sheet.")
    sys.exit()

template_position = template_vp.GetBoxCenter()

# Step 3 - Select target sheets to align views
form = SheetSelectionWindow(all_sheets, multiselect=True, title="Select Sheets to Align Views")
if not form.ShowDialog() or not form.selected_sheets:
    sys.exit()
target_sheets = form.selected_sheets

# Step 4 - Align views on selected sheets
aligned = 0
trans = DB.Transaction(doc, "Align Views on Sheets")
try:
    trans.Start()
    for sheet in target_sheets:
        for vp_id in sheet.GetAllViewports():
            vp = doc.GetElement(vp_id)
            vp.SetBoxCenter(template_position)
            aligned += 1
    trans.Commit()
except Exception as e:
    trans.RollBack()
    TaskDialog.Show("Error", "Failed to align views: " + str(e))
    sys.exit()
finally:
    if trans.HasStarted() and not trans.HasEnded():
        trans.RollBack()

TaskDialog.Show("Alignment Complete", "Aligned %d view(s) to match template position." % aligned)