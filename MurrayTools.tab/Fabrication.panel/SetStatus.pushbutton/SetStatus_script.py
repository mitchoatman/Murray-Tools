# -*- coding: utf-8 -*-
# Imports
import Autodesk
from Autodesk.Revit.DB import FabricationConfiguration, ElementId, Transaction
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
import clr, sys

clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('System')

from System.Windows import (
    Application, Window, Thickness,
    HorizontalAlignment, VerticalAlignment,
    ResizeMode, WindowStartupLocation,
    GridLength, GridUnitType
)
from System.Windows.Controls import (
    Button, Grid, RowDefinition, ColumnDefinition,
    Label, StackPanel, ScrollViewer,
    Orientation, ScrollBarVisibility,
    ComboBox, ListBox, SelectionMode
)
from System.Windows.Media import Brushes, FontFamily

# Get document and UIDocument
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Get fabrication configuration and status list
config = FabricationConfiguration.GetFabricationConfiguration(doc)
status_names = []

if config:
    status_ids = config.GetAllPartStatuses()
    if status_ids and status_ids.Count > 0:
        for status_id in status_ids:
            name = config.GetPartStatusDescription(status_id)
            status_names.append(name)
    else:
        status_names = ["None"]
else:
    status_names = ["None"]

# Get all unique STRATUS Status values from elements in the current view
view = doc.ActiveView
collector = Autodesk.Revit.DB.FilteredElementCollector(doc, view.Id)
elements = collector.WhereElementIsNotElementType().ToElements()
stratus_statuses = set()

for elem in elements:
    param = elem.LookupParameter("STRATUS Status")
    if param and param.StorageType == Autodesk.Revit.DB.StorageType.String and param.HasValue:
        stratus_statuses.add(param.AsString())

stratus_statuses = list(stratus_statuses)
if not stratus_statuses:
    stratus_statuses = ["None"]


# Custom selection filter
class StatusSelectionFilter(ISelectionFilter):
    def __init__(self, excluded_statuses):
        self.excluded_statuses = excluded_statuses

    def AllowElement(self, element):
        param = element.LookupParameter("STRATUS Status")
        if param and param.StorageType == Autodesk.Revit.DB.StorageType.String and param.HasValue:
            return param.AsString() not in self.excluded_statuses
        return True

    def AllowReference(self, reference, point):
        return True


# WPF Form for selecting statuses to exclude and apply
class StatusSelectionForm(object):
    def __init__(self, stratus_statuses, status_names):
        self.selected_excluded_statuses = []
        self.selected_status_to_apply = None
        self.stratus_statuses = sorted(stratus_statuses)
        self.status_names = sorted(status_names)
        self.InitializeComponents()

    def InitializeComponents(self):
        self._window = Window()
        self._window.Title = "Set STRATUS Status"
        self._window.Width = 400
        self._window.Height = 350
        self._window.MinWidth = self._window.Width
        self._window.MinHeight = self._window.Height
        self._window.ResizeMode = ResizeMode.NoResize
        self._window.WindowStartupLocation = WindowStartupLocation.CenterScreen

        grid = Grid()
        grid.Margin = Thickness(10)
        grid.VerticalAlignment = VerticalAlignment.Stretch
        grid.HorizontalAlignment = HorizontalAlignment.Stretch

        # Define grid rows
        for i in range(5):
            row = RowDefinition()
            if i == 1:
                row.Height = GridLength(0.5, GridUnitType.Star)
            else:
                row.Height = GridLength.Auto
            grid.RowDefinitions.Add(row)
        grid.ColumnDefinitions.Add(ColumnDefinition())

        # Label for exclude statuses
        exclude_label = Label()
        exclude_label.Content = "Select STRATUS Statuses to Exclude:"
        exclude_label.FontFamily = FontFamily("Arial")
        exclude_label.FontSize = 14
        exclude_label.Margin = Thickness(0, 0, 0, 5)
        Grid.SetRow(exclude_label, 0)
        grid.Children.Add(exclude_label)

        # ScrollViewer for ListBox
        scroll_viewer = ScrollViewer()
        scroll_viewer.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll_viewer.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
        scroll_viewer.VerticalAlignment = VerticalAlignment.Stretch
        scroll_viewer.Margin = Thickness(0, 0, 0, 10)

        # ListBox for exclude statuses
        self.exclude_listbox = ListBox()
        self.exclude_listbox.SelectionMode = SelectionMode.Extended
        self.exclude_listbox.FontFamily = FontFamily("Arial")
        self.exclude_listbox.FontSize = 12
        for status in self.stratus_statuses:
            self.exclude_listbox.Items.Add(status)
        self.exclude_listbox.SelectionChanged += self.exclude_selection_changed
        scroll_viewer.Content = self.exclude_listbox
        Grid.SetRow(scroll_viewer, 1)
        grid.Children.Add(scroll_viewer)

        # Label for apply status
        apply_label = Label()
        apply_label.Content = "Select Fabrication Status to Apply:"
        apply_label.FontFamily = FontFamily("Arial")
        apply_label.FontSize = 14
        apply_label.Margin = Thickness(0, 5, 0, 5)
        Grid.SetRow(apply_label, 2)
        grid.Children.Add(apply_label)

        # ComboBox for apply status
        self.apply_combobox = ComboBox()
        self.apply_combobox.Height = 20
        self.apply_combobox.FontFamily = FontFamily("Arial")
        self.apply_combobox.FontSize = 12
        self.apply_combobox.Margin = Thickness(0, 0, 0, 10)
        for status in self.status_names:
            self.apply_combobox.Items.Add(status)
        self.apply_combobox.SelectionChanged += self.apply_selection_changed
        Grid.SetRow(self.apply_combobox, 3)
        grid.Children.Add(self.apply_combobox)

        # Button panel
        button_container = Grid()
        button_container.HorizontalAlignment = HorizontalAlignment.Stretch
        button_container.VerticalAlignment = VerticalAlignment.Bottom

        button_panel = StackPanel()
        button_panel.HorizontalAlignment = HorizontalAlignment.Center
        button_panel.VerticalAlignment = VerticalAlignment.Bottom
        button_panel.Orientation = Orientation.Horizontal
        button_panel.Margin = Thickness(0, 10, 0, 10)

        # Set Status button
        set_button = Button()
        set_button.Content = "Set Status"
        set_button.FontFamily = FontFamily("Arial")
        set_button.FontSize = 12
        set_button.Height = 25
        set_button.Width = 80
        set_button.Margin = Thickness(5, 0, 5, 0)
        set_button.Click += self.set_clicked
        button_panel.Children.Add(set_button)

        button_container.Children.Add(button_panel)
        Grid.SetRow(button_container, 4)
        grid.Children.Add(button_container)

        self._window.Content = grid

    def exclude_selection_changed(self, sender, args):
        self.selected_excluded_statuses = [item for item in self.exclude_listbox.SelectedItems]

    def apply_selection_changed(self, sender, args):
        self.selected_status_to_apply = self.apply_combobox.SelectedItem

    def set_clicked(self, sender, args):
        self._window.Close()

    def ShowDialog(self):
        self._window.ShowDialog()
        return self.selected_excluded_statuses, self.selected_status_to_apply


# -------- Script Execution --------
# 1. Show dialog
form = StatusSelectionForm(stratus_statuses, status_names)
excluded_statuses, selected_status = form.ShowDialog()

if not selected_status:
    TaskDialog.Show("Error", "No status selected to apply.")
    sys.exit()

# 2. Prompt user to select elements with filter
try:
    selection_filter = StatusSelectionFilter(excluded_statuses)
    selected_elements = uidoc.Selection.PickObjects(
        ObjectType.Element,
        selection_filter,
        "Select elements to assign status"
    )
except Exception as e:
    TaskDialog.Show("Error", "Selection canceled or failed: {}".format(str(e)))
    sys.exit()

if not selected_elements:
    TaskDialog.Show("Error", "No elements selected.")
    sys.exit()

# 3. Apply selected status
t = Transaction(doc, "Set STRATUS Status")
t.Start()
try:
    for ref in selected_elements:
        elem = doc.GetElement(ref.ElementId)
        param = elem.LookupParameter("STRATUS Status")
        if param and param.StorageType == Autodesk.Revit.DB.StorageType.String:
            if not excluded_statuses or param.AsString() not in excluded_statuses:
                param.Set(selected_status)
    t.Commit()
except Exception as e:
    t.RollBack()
    TaskDialog.Show("Error", "Failed to set STRATUS Status: {}".format(str(e)))
