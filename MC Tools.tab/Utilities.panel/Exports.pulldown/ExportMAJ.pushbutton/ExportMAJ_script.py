# -*- coding: UTF-8 -*-
import Autodesk
import clr
import os
import sys

clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('System')
clr.AddReference("System.Core")
clr.AddReference('System.Windows.Forms')

from Autodesk.Revit.DB import FabricationConfiguration, ElementId, Transaction
from Autodesk.Revit.DB.Fabrication import FabricationSaveJobOptions
from Autodesk.Revit.DB.FabricationPart import SaveAsFabricationJob
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from System.Collections.Generic import List, HashSet
from System.Windows.Forms import SaveFileDialog, DialogResult

from System.Windows import Thickness, HorizontalAlignment, VerticalAlignment, ResizeMode, WindowStartupLocation, GridLength, GridUnitType, Window
from System.Windows.Controls import Button, TextBox, CheckBox, RadioButton, Grid, RowDefinition, ColumnDefinition, Label, StackPanel, ScrollViewer, ScrollBarVisibility
from System.Windows.Controls.Primitives import UniformGrid
from System.Windows.Media import FontFamily, Brushes

# Get document and UIDocument
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


class SearchableSelectionForm(object):
    def __init__(self, items, title, label_text, multiselect=True, preselected=None):
        self.all_items = sorted([x for x in items if x])
        self.title = title
        self.label_text = label_text
        self.multiselect = multiselect
        self.selected_items = set(preselected) if preselected else set()
        self.result = None
        self.controls = []
        self.check_all_state = False
        self.InitializeComponents()

    def InitializeComponents(self):
        self._window = Window()
        self._window.Title = self.title
        self._window.Width = 425
        self._window.Height = 500
        self._window.MinWidth = 425
        self._window.MinHeight = 500
        self._window.ResizeMode = ResizeMode.NoResize
        self._window.WindowStartupLocation = WindowStartupLocation.CenterScreen

        grid = Grid()
        grid.Margin = Thickness(8)

        for i in range(4):
            row = RowDefinition()
            if i == 2:
                row.Height = GridLength(1, GridUnitType.Star)
            else:
                row.Height = GridLength.Auto
            grid.RowDefinitions.Add(row)

        grid.ColumnDefinitions.Add(ColumnDefinition())

        self.label = Label()
        self.label.Content = self.label_text
        self.label.FontFamily = FontFamily("Arial")
        self.label.FontSize = 16
        self.label.Margin = Thickness(0, 0, 0, 5)
        Grid.SetRow(self.label, 0)
        grid.Children.Add(self.label)

        self.search_box = TextBox()
        self.search_box.Height = 24
        self.search_box.Margin = Thickness(0, 0, 0, 5)
        self.search_box.FontFamily = FontFamily("Arial")
        self.search_box.FontSize = 12
        self.search_box.TextChanged += self.search_changed
        Grid.SetRow(self.search_box, 1)
        grid.Children.Add(self.search_box)

        self.item_panel = StackPanel()
        scroll_viewer = ScrollViewer()
        scroll_viewer.Content = self.item_panel
        scroll_viewer.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll_viewer.Margin = Thickness(0, 0, 0, 5)
        Grid.SetRow(scroll_viewer, 2)
        grid.Children.Add(scroll_viewer)

        button_container = Grid()
        button_panel = UniformGrid()
        button_panel.Columns = 2

        button_panel.HorizontalAlignment = HorizontalAlignment.Center
        button_panel.VerticalAlignment = VerticalAlignment.Bottom
        button_panel.Margin = Thickness(0, 5, 0, 0)

        self.confirm_button = Button()
        self.confirm_button.Content = "Confirm"
        self.confirm_button.FontFamily = FontFamily("Arial")
        self.confirm_button.FontSize = 12
        self.confirm_button.Height = 28
        self.confirm_button.Margin = Thickness(5, 0, 5, 0)
        self.confirm_button.Click += self.confirm_clicked
        button_panel.Children.Add(self.confirm_button)

        self.cancel_button = Button()
        self.cancel_button.Content = "None"
        self.cancel_button.Background = Brushes.LightGray
        self.cancel_button.FontFamily = FontFamily("Arial")
        self.cancel_button.FontSize = 12
        self.cancel_button.Height = 28
        self.cancel_button.Margin = Thickness(5, 0, 5, 0)
        self.cancel_button.Click += self.cancel_clicked
        button_panel.Children.Add(self.cancel_button)

        button_container.Children.Add(button_panel)
        Grid.SetRow(button_container, 3)
        grid.Children.Add(button_container)

        self._window.Content = grid
        self.update_items(self.all_items)

    def update_items(self, items):
        self.item_panel.Children.Clear()
        self.controls = []

        for item in items:
            if self.multiselect:
                control = CheckBox()
                control.Checked += self.item_checked
                control.Unchecked += self.item_unchecked
            else:
                control = RadioButton()
                control.GroupName = "SelectionGroup"
                control.Checked += self.radio_checked

            control.Content = item
            control.FontFamily = FontFamily("Arial")
            control.FontSize = 12
            control.Margin = Thickness(2)
            control.IsChecked = item in self.selected_items

            self.item_panel.Children.Add(control)
            self.controls.append(control)

    def search_changed(self, sender, args):
        search_text = self.search_box.Text.lower().strip()
        if search_text:
            filtered_items = [x for x in self.all_items if search_text in x.lower()]
        else:
            filtered_items = self.all_items
        self.update_items(filtered_items)

    def item_checked(self, sender, args):
        self.selected_items.add(sender.Content)

    def item_unchecked(self, sender, args):
        if sender.Content in self.selected_items:
            self.selected_items.remove(sender.Content)

    def radio_checked(self, sender, args):
        self.selected_items = set([sender.Content])

    def confirm_clicked(self, sender, args):
        if self.multiselect:
            self.result = sorted(list(self.selected_items))
        else:
            self.result = list(self.selected_items)[0] if self.selected_items else None
        self._window.Close()

    def cancel_clicked(self, sender, args):
        self.result = None
        self._window.Close()

    def ShowDialog(self):
        self._window.ShowDialog()
        return self.result


# Get fabrication configuration and status list
config = FabricationConfiguration.GetFabricationConfiguration(doc)
status_names = []

if config:
    status_ids = config.GetAllPartStatuses()
    if status_ids and status_ids.Count > 0:
        for status_id in status_ids:
            name = config.GetPartStatusDescription(status_id)
            if name:
                status_names.append(name)
    else:
        print "No statuses found in the fabrication configuration."
        status_names = ["None"]
else:
    print "No fabrication configuration found in this project."
    status_names = ["None"]

status_names = sorted(list(set(status_names)))

# Get all unique STRATUS Status values from elements in the current view
view = doc.ActiveView
collector = Autodesk.Revit.DB.FilteredElementCollector(doc, view.Id)
elements = collector.WhereElementIsNotElementType().ToElements()
stratus_statuses = set()

for elem in elements:
    param = elem.LookupParameter("STRATUS Status")
    if param and param.StorageType == Autodesk.Revit.DB.StorageType.String and param.HasValue:
        val = param.AsString()
        if val:
            stratus_statuses.add(val)

stratus_statuses = sorted(list(stratus_statuses))
if not stratus_statuses:
    stratus_statuses = ["None"]

# Searchable dialog to select statuses to exclude
excluded_statuses = SearchableSelectionForm(
    stratus_statuses,
    "Select Status to Exclude",
    "Choose status to exclude in selection:",
    multiselect=True
).ShowDialog()

if not excluded_statuses:
    print "No statuses excluded. All elements can be selected."
    excluded_statuses = []


# Custom selection filter to exclude elements with selected statuses
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


# Prompt user to select elements with the custom filter
try:
    selection_filter = StatusSelectionFilter(excluded_statuses)
    selected_elements = uidoc.Selection.PickObjects(
        ObjectType.Element,
        selection_filter,
        "Select elements to export and assign status (excluded statuses filtered)"
    )
except Exception as e:
    print "Selection canceled or failed: %s" % str(e)
    sys.exit()

if not selected_elements:
    print "No elements selected. Exiting."
    sys.exit()

element_ids = [elem.ElementId for elem in selected_elements]
id_set = HashSet[ElementId]()
for id in element_ids:
    id_set.Add(id)

# Searchable single-select dialog for fabrication status
selected_status = SearchableSelectionForm(
    status_names,
    "Select Status",
    "Choose status to apply to exported:",
    multiselect=False
).ShowDialog()

if not selected_status:
    selected_status = None

# MAJ export logic
options = FabricationSaveJobOptions()
folder_name = "C:\\Temp"
filepath = os.path.join(folder_name, "Ribbon_Exports.txt")

if not os.path.exists(folder_name):
    os.makedirs(folder_name)

default_desktop_path = os.path.expandvars("%USERPROFILE%\\Desktop")
if os.path.exists(filepath):
    with open(filepath, 'r') as f:
        last_save_path = f.read().strip()
    if not os.path.exists(last_save_path):
        last_save_path = default_desktop_path
else:
    last_save_path = default_desktop_path

save_dialog = SaveFileDialog()
save_dialog.Filter = "Fabrication Job Files (*.maj)|*.maj"
save_dialog.DefaultExt = "maj"
save_dialog.InitialDirectory = last_save_path
save_dialog.FileName = doc.Title
save_dialog.Title = "Save MAJ File"

result = save_dialog.ShowDialog()

if result == DialogResult.OK:
    file_path = save_dialog.FileName
    folder_path = os.path.dirname(file_path)

    try:
        SaveAsFabricationJob(doc, id_set, file_path, options)

        with open(filepath, 'w') as f:
            f.write(folder_path)

        if selected_status:
            t = Transaction(doc, "Set STRATUS Status")
            t.Start()
            try:
                for ref in selected_elements:
                    elem = doc.GetElement(ref.ElementId)
                    param = elem.LookupParameter("STRATUS Status")
                    if param and param.StorageType == Autodesk.Revit.DB.StorageType.String:
                        param.Set(selected_status)
                t.Commit()
            except Exception as e:
                t.RollBack()
                print "Failed to set STRATUS Status: %s" % str(e)

    except Exception as e:
        print "Export failed: %s" % str(e)
else:
    print "Fabrication job saving canceled."