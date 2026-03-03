# -*- coding: utf-8 -*-
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Collections')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from System.Windows import Application, Window, WindowStartupLocation, WindowStyle, Visibility, GridLength, HorizontalAlignment, VerticalAlignment, GridUnitType
from System.Windows.Controls import Label, ComboBox, Button, ListBox, CheckBox, TextBox, Grid, RowDefinition, ColumnDefinition, SelectionMode, StackPanel, Orientation, ListBoxItem
from System.Windows import Thickness
from System.Windows.Media import Brushes
from System.Collections.Generic import List
from System.Windows.Threading import DispatcherFrame, Dispatcher
from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, FabricationConfiguration, Transaction, TemporaryViewMode, ParameterValueProvider, FilterStringRule, ElementParameterFilter, FilterStringContains
from Autodesk.Revit.UI import UIDocument, TaskDialog, TaskDialogCommonButtons
from Parameters.Add_SharedParameters import Shared_Params
from Parameters.Get_Set_Params import (get_parameter_value_by_name_AsString,
                                     get_parameter_value_by_name_AsValueString,
                                     get_parameter_value_by_name_AsInteger)
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
Config = FabricationConfiguration.GetFabricationConfiguration(doc)
Shared_Params()
class ValueItem(object):
    def __init__(self, property, value):
        self.Property = property
        self.Value = value
    def __str__(self):
        return u"{}: {}".format(self.Property, self.Value)
class RemoveFilterDialog(Window):
    def __init__(self, filter_options, filter_keys):
        self.filter_options = filter_options
        self.filter_keys = filter_keys
        self.selected_filter = None
        self.InitializeComponents()
    def InitializeComponents(self):
        self.Title = "Remove Filter"
        self.Width = 300
        self.Height = 140
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.ResizeMode = 0 # CanMinimize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Topmost = True # Ensure dialog appears on top
        # Create grid for layout
        grid = Grid()
        self.Content = grid
        # Define rows
        row_definitions = [
            RowDefinition(Height=GridLength.Auto), # Label and ComboBox
            RowDefinition(Height=GridLength.Auto) # Buttons
        ]
        for row in row_definitions:
            grid.RowDefinitions.Add(row)
        # Label
        self.label = Label()
        self.label.Content = "Select filter to remove:"
        self.label.Margin = Thickness(10, 5, 0, 0)
        self.label.Visibility = Visibility.Visible
        Grid.SetRow(self.label, 0)
        grid.Children.Add(self.label)
        # ComboBox
        self.filter_combo = ComboBox()
        self.filter_combo.Margin = Thickness(10, 30, 10, 0)
        self.filter_combo.Width = 260
        self.filter_combo.Height = 20
        self.filter_combo.IsDropDownOpen = False
        self.filter_combo.Visibility = Visibility.Visible
        for option in self.filter_options:
            self.filter_combo.Items.Add(option)
        if self.filter_options:
            self.filter_combo.SelectedIndex = 0
        Grid.SetRow(self.filter_combo, 0)
        grid.Children.Add(self.filter_combo)
        # Button panel for Remove, Remove All, and Cancel buttons
        button_panel = StackPanel()
        button_panel.Orientation = Orientation.Horizontal
        button_panel.HorizontalAlignment = HorizontalAlignment.Center
        button_panel.Margin = Thickness(0, 10, 0, 10) # Reduced top margin for tighter spacing
        Grid.SetRow(button_panel, 1)
        grid.Children.Add(button_panel)
        # Remove Button
        self.ok_button = Button()
        self.ok_button.Content = "Remove"
        self.ok_button.Width = 80
        self.ok_button.Height = 25
        self.ok_button.Margin = Thickness(0, 0, 5, 0)
        self.ok_button.Visibility = Visibility.Visible
        self.ok_button.Click += self.ok_clicked
        button_panel.Children.Add(self.ok_button)
        # Remove All Button
        self.remove_all_button = Button()
        self.remove_all_button.Content = "Remove All"
        self.remove_all_button.Width = 80
        self.remove_all_button.Height = 25
        self.remove_all_button.Margin = Thickness(5, 0, 5, 0)
        self.remove_all_button.Visibility = Visibility.Visible
        self.remove_all_button.Click += self.remove_all_clicked
        button_panel.Children.Add(self.remove_all_button)
        # Cancel Button
        self.cancel_button = Button()
        self.cancel_button.Content = "Cancel"
        self.cancel_button.Width = 80
        self.cancel_button.Height = 25
        self.cancel_button.Margin = Thickness(5, 0, 0, 0)
        self.cancel_button.Visibility = Visibility.Visible
        self.cancel_button.Click += self.cancel_clicked
        button_panel.Children.Add(self.cancel_button)
    def ok_clicked(self, sender, args):
        if self.filter_combo.SelectedIndex >= 0:
            self.selected_filter = self.filter_keys[self.filter_combo.SelectedIndex]
        self.DialogResult = True
        self.Close()
    def remove_all_clicked(self, sender, args):
        self.selected_filter = None # Special value to indicate remove all
        self.DialogResult = True
        self.Close()
    def cancel_clicked(self, sender, args):
        self.DialogResult = False
        self.Close()
class MultiPropertyFilterForm(Window):
    def __init__(self, property_options, fab_elements, all_elements):
        self.property_options = property_options
        self.fab_elements = fab_elements
        self.all_elements = all_elements
        self.selected_filters = {}
        self.InitializeComponents()
        if self.property_options:
            first_property = self.property_combo.Items[0]
            if first_property in self.property_options:
                self.property_combo.SelectedItem = first_property
            self.update_values_list(None, None)
        self.update_filter_display()
    def InitializeComponents(self):
        self.Title = "Multi-Property Filter"
        self.Width = 550
        self.Height = 620
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.ResizeMode = 0 # CanMinimize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Topmost = True # Keep dialog always on top
        # Create grid with row and column definitions
        grid = Grid()
        self.Content = grid
        # Define rows with GridLength objects
        row_definitions = [
            RowDefinition(Height=GridLength(40)), # Select Property
            RowDefinition(Height=GridLength.Auto), # Search
            RowDefinition(Height=GridLength.Auto), # Select Values
            RowDefinition(Height=GridLength(260)), # ListBox
            RowDefinition(Height=GridLength.Auto), # Add/Remove Filter buttons
            RowDefinition(Height=GridLength(140)), # Filter display
            RowDefinition(Height=GridLength.Auto) # Button row
        ]
        for row in row_definitions:
            grid.RowDefinitions.Add(row)
        # Define columns for all components
        column_definitions = [
            ColumnDefinition(Width=GridLength(1, GridUnitType.Star)), # Left padding
            ColumnDefinition(Width=GridLength.Auto), # Main content
            ColumnDefinition(Width=GridLength(1, GridUnitType.Star)) # Right padding
        ]
        for col in column_definitions:
            grid.ColumnDefinitions.Add(col)
        # Property selection label
        self.property_label = Label()
        self.property_label.Content = "Select Property:"
        self.property_label.Margin = Thickness(10, 0, 10, 0)
        self.property_label.Visibility = Visibility.Visible
        self.property_label.HorizontalAlignment = HorizontalAlignment.Left
        Grid.SetRow(self.property_label, 0)
        Grid.SetColumn(self.property_label, 1)
        grid.Children.Add(self.property_label)
        # Property ComboBox
        self.property_combo = ComboBox()
        self.property_combo.Width = 160
        self.property_combo.Height = 22
        self.property_combo.Margin = Thickness(120, 0, 10, 0)
        self.property_combo.IsDropDownOpen = False
        self.property_combo.Visibility = Visibility.Visible
        self.property_combo.HorizontalAlignment = HorizontalAlignment.Left
        properties = sorted(self.property_options.keys()) # Sort properties alphabetically
        for prop in properties:
            self.property_combo.Items.Add(prop)
        self.property_combo.SelectionChanged += self.update_values_list
        Grid.SetRow(self.property_combo, 0)
        Grid.SetColumn(self.property_combo, 1)
        grid.Children.Add(self.property_combo)
        # AND/OR toggle
        self.logic_check = CheckBox()
        self.logic_check.Content = "AND logic (unchecked = OR)"
        self.logic_check.Width = 230
        self.logic_check.Margin = Thickness(10, 5, 10, 0)
        self.logic_check.Visibility = Visibility.Visible
        self.logic_check.HorizontalAlignment = HorizontalAlignment.Center
        Grid.SetRow(self.logic_check, 4)
        Grid.SetColumn(self.logic_check, 1)
        grid.Children.Add(self.logic_check)
        # Search label
        self.search_label = Label()
        self.search_label.Content = "Search:"
        self.search_label.Margin = Thickness(10, 0, 10, 0)
        self.search_label.Visibility = Visibility.Visible
        self.search_label.HorizontalAlignment = HorizontalAlignment.Left
        Grid.SetRow(self.search_label, 1)
        Grid.SetColumn(self.search_label, 1)
        grid.Children.Add(self.search_label)
        # Search bar
        self.search_box = TextBox()
        self.search_box.Width = 300
        self.search_box.Height = 20
        self.search_box.Margin = Thickness(75, 0, 10, 0)
        self.search_box.Visibility = Visibility.Visible
        self.search_box.HorizontalAlignment = HorizontalAlignment.Left
        self.search_box.TextChanged += self.update_values_list
        Grid.SetRow(self.search_box, 1)
        Grid.SetColumn(self.search_box, 1)
        grid.Children.Add(self.search_box)
        self.search_box.Focus()
        # Values list label
        self.values_label = Label()
        self.values_label.Content = "Select Values:"
        self.values_label.Margin = Thickness(10, 25, 10, 0)
        self.values_label.Visibility = Visibility.Visible
        self.values_label.HorizontalAlignment = HorizontalAlignment.Left
        Grid.SetRow(self.values_label, 2)
        Grid.SetColumn(self.search_label, 1)
        grid.Children.Add(self.values_label)
        # Values list
        self.values_list = ListBox()
        self.values_list.Width = 500
        self.values_list.Height = 250
        self.values_list.Margin = Thickness(10, 0, 10, 0)
        self.values_list.SelectionMode = SelectionMode.Extended
        self.values_list.Visibility = Visibility.Visible
        self.values_list.HorizontalAlignment = HorizontalAlignment.Center
        self.values_list.MouseDoubleClick += self.add_filter
        Grid.SetRow(self.values_list, 3)
        Grid.SetColumn(self.values_list, 1)
        grid.Children.Add(self.values_list)
        # Add Filter button
        self.add_button = Button()
        self.add_button.Content = "Add Filter"
        self.add_button.Width = 80
        self.add_button.Height = 25
        self.add_button.Margin = Thickness(10, 20, 10, 0)
        self.add_button.Visibility = Visibility.Visible
        self.add_button.ToolTip = "You can also double click on properties"
        self.add_button.HorizontalAlignment = HorizontalAlignment.Right
        self.add_button.Click += self.add_filter
        Grid.SetRow(self.add_button, 2)
        Grid.SetColumn(self.add_button, 1)
        grid.Children.Add(self.add_button)
        # Filter management button
        self.filter_button = Button()
        self.filter_button.Width = 500
        self.filter_button.Height = 130
        self.filter_button.Margin = Thickness(10, 10, 10, 0)
        self.filter_button.Visibility = Visibility.Visible
        self.filter_button.ToolTip = "Click here to modify Filters"
        self.filter_button.HorizontalContentAlignment = HorizontalAlignment.Left
        self.filter_button.VerticalContentAlignment = VerticalAlignment.Top
        self.filter_button.HorizontalAlignment = HorizontalAlignment.Center
        self.filter_button.Click += self.remove_filter
        Grid.SetRow(self.filter_button, 5)
        Grid.SetColumn(self.filter_button, 1)
        grid.Children.Add(self.filter_button)
        # Button panel for Update Data, Reset View, Isolate, Select, and Cancel buttons
        button_panel = StackPanel()
        button_panel.Orientation = Orientation.Horizontal
        button_panel.HorizontalAlignment = HorizontalAlignment.Center
        button_panel.Margin = Thickness(0, 15, 0, 0)
        Grid.SetRow(button_panel, 6)
        Grid.SetColumn(button_panel, 1)
        grid.Children.Add(button_panel)
        # Update Data Button
        self.reset_filter_button = Button()
        self.reset_filter_button.Content = "Update Data"
        self.reset_filter_button.Width = 80
        self.reset_filter_button.Height = 25
        self.reset_filter_button.Margin = Thickness(0, 0, 5, 0)
        self.reset_filter_button.Visibility = Visibility.Visible
        self.reset_filter_button.ToolTip = "Refreshes filter data based on current view content"
        self.reset_filter_button.Click += self.reset_filter_clicked
        button_panel.Children.Add(self.reset_filter_button)
        # Reset View Button
        self.reset_button = Button()
        self.reset_button.Content = "Reset View"
        self.reset_button.Width = 80
        self.reset_button.Height = 25
        self.reset_button.Foreground = Brushes.Black
        self.reset_button.Background = Brushes.Red
        self.reset_button.Margin = Thickness(5, 0, 5, 0)
        self.reset_button.Visibility = Visibility.Visible
        self.reset_button.Click += self.reset_clicked
        button_panel.Children.Add(self.reset_button)
        # Isolate Button
        self.isolate_button = Button()
        self.isolate_button.Content = "Isolate"
        self.isolate_button.Width = 80
        self.isolate_button.Height = 25
        self.isolate_button.Margin = Thickness(5, 0, 5, 0)
        self.isolate_button.Visibility = Visibility.Visible
        self.isolate_button.Click += self.isolate_clicked
        button_panel.Children.Add(self.isolate_button)
        # Select Button
        self.select_button = Button()
        self.select_button.Content = "Select"
        self.select_button.Width = 80
        self.select_button.Height = 25
        self.select_button.Margin = Thickness(5, 0, 5, 0)
        self.select_button.Visibility = Visibility.Visible
        self.select_button.Click += self.select_clicked
        button_panel.Children.Add(self.select_button)
        # Cancel Button
        self.cancel_button = Button()
        self.cancel_button.Content = "Close"
        self.cancel_button.Width = 80
        self.cancel_button.Height = 25
        self.cancel_button.Margin = Thickness(5, 0, 0, 0)
        self.cancel_button.Visibility = Visibility.Visible
        self.cancel_button.Click += self.cancel_clicked
        button_panel.Children.Add(self.cancel_button)
    def exit_frame(self, sender, e):
        self.frame.Continue = False
    def reset_filter_clicked(self, sender, args):
        try:
            # Refresh elements based on current view
            preselection = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()]
            self.fab_elements = preselection if preselection else FilteredElementCollector(doc, curview.Id) \
                .OfClass(DB.FabricationPart) \
                .WhereElementIsNotElementType() \
                .ToElements()
            self.all_elements = preselection if preselection else FilteredElementCollector(doc, curview.Id) \
                .WhereElementIsNotElementType() \
                .ToElements()
           
            # Rebuild property options
            self.property_options = {}
            for prop in ['CID', 'ServiceType', 'Service Name', 'Service Abbreviation', 'Size',
                        'STRATUS Assembly', 'Line Number', 'STRATUS Status', 'Reference Level',
                        'Item Number', 'Bundle Number', 'REF BS Designation', 'REF Line Number',
                        'Specification', 'Hanger Rod Size', 'Valve Number', 'Beam Hanger', 'Product Entry']:
                values = set(filter(None, [get_property_value(elem, prop, debug=False) for elem in self.fab_elements]))
                if values:
                    self.property_options[prop] = sorted(values)
            for prop in ['Name', 'Comments', 'Category']:
                if self.all_elements:
                    values = set(filter(None, [get_property_value(elem, prop, debug=False) for elem in self.all_elements]))
                    if values:
                        self.property_options[prop] = sorted(values)
            # Update ComboBox
            self.property_combo.Items.Clear()
            properties = sorted(self.property_options.keys()) # Sort properties alphabetically
            for prop in properties:
                self.property_combo.Items.Add(prop)
            if properties:
                self.property_combo.SelectedIndex = 0
                # Explicitly refresh the values list for the selected property
                self.property_combo.SelectedItem = properties[0]
                self.values_list.Items.Clear()
                self.search_box.Text = ""
                values = self.property_options.get(properties[0], [])
                if values:
                    for value in values:
                        item = ListBoxItem()
                        value_item = ValueItem(properties[0], value)
                        item.Content = str(value_item)
                        item.Tag = value_item
                        self.values_list.Items.Add(item)
           
            # Clear search box
            self.search_box.Text = ""
           
        except Exception as e:
            dialog = TaskDialog("Error")
            dialog.MainInstruction = "Update Data Error: {}".format(str(e))
            dialog.CommonButtons = TaskDialogCommonButtons.Ok
            dialog.Show()
    def update_values_list(self, sender, args):
        search_term = self.search_box.Text.lower()
        selected_property = self.property_combo.SelectedItem
        self.values_list.Items.Clear()
        if search_term:
            matching = []
            for prop in self.property_options:
                values = self.property_options.get(prop, [])
                filtered_values = [v for v in values if search_term in str(v).lower()]
                for v in filtered_values:
                    matching.append(ValueItem(prop, v))
            matching.sort(key=lambda x: (x.Property, x.Value))
            for value_item in matching:
                item = ListBoxItem()
                item.Content = str(value_item)
                item.Tag = value_item
                self.values_list.Items.Add(item)
        else:
            if selected_property:
                values = self.property_options.get(selected_property, [])
                for v in sorted(values):
                    item = ListBoxItem()
                    value_item = ValueItem(selected_property, v)
                    item.Content = str(value_item)
                    item.Tag = value_item
                    self.values_list.Items.Add(item)
    def add_filter(self, sender, args):
        selected_items = self.values_list.SelectedItems
        if selected_items:
            from collections import defaultdict
            filters_to_add = defaultdict(list)
            for item in selected_items:
                value_item = item.Tag # Retrieve ValueItem from Tag
                filters_to_add[value_item.Property].append(value_item.Value)
            for prop, vals in filters_to_add.items():
                if prop not in self.selected_filters:
                    self.selected_filters[prop] = []
                self.selected_filters[prop].append((vals, self.logic_check.IsChecked))
            self.update_filter_display()
        else:
            dialog = TaskDialog("Warning")
            dialog.MainInstruction = "Please select at least one value."
            dialog.CommonButtons = TaskDialogCommonButtons.Ok
            dialog.Show()
    def remove_filter(self, sender, args):
        if not self.selected_filters:
            return # Do nothing if no filters exist
       
        filter_options = []
        filter_keys = []
        for prop, filter_list in self.selected_filters.items():
            for values, is_and in filter_list:
                mode = "AND" if is_and else "OR"
                desc = "%s (%s): %s" % (prop, mode, ", ".join(str(v) for v in values))
                filter_options.append(desc)
                filter_keys.append((prop, values))
       
        if not filter_options:
            dialog = TaskDialog("Information")
            dialog.MainInstruction = "No filters to remove."
            dialog.CommonButtons = TaskDialogCommonButtons.Ok
            dialog.Show()
            return
       
        dialog = RemoveFilterDialog(filter_options, filter_keys)
        if dialog.ShowDialog() == True:
            if dialog.selected_filter is None:
                # Remove all filters
                self.selected_filters = {}
            elif dialog.selected_filter:
                prop, values = dialog.selected_filter
                for i, (v, is_and) in enumerate(self.selected_filters[prop]):
                    if v == values:
                        del self.selected_filters[prop][i]
                        break
                if not self.selected_filters[prop]:
                    del self.selected_filters[prop]
            self.update_filter_display()
    def update_filter_display(self, sender=None, args=None):
        if not self.selected_filters:
            self.filter_button.Content = "No Filters Yet..."
        else:
            filter_text = "Filters (Click here to modify):\n"
            for prop, filter_list in self.selected_filters.items():
                filter_text += "%s:\n " % prop
                conditions = []
                for values, is_and in filter_list:
                    mode = "AND" if is_and else "OR"
                    condition = "[%s: %s]" % (mode, ", ".join(str(v) for v in values))
                    conditions.append(condition)
                filter_text += " ".join(conditions) + "\n"
            self.filter_button.Content = filter_text.strip()
    def reset_clicked(self, sender, args):
        try:
            t = Transaction(doc, "Reset Temporary Hide/Isolate")
            t.Start()
            curview.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate)
            t.Commit()
        except Exception as e:
            dialog = TaskDialog("Error")
            dialog.MainInstruction = "Reset Error: {}".format(str(e))
            dialog.CommonButtons = TaskDialogCommonButtons.Ok
            dialog.Show()
    def isolate_clicked(self, sender, args):
        if not self.selected_filters:
            dialog = TaskDialog("Warning")
            dialog.MainInstruction = "No filters selected to isolate."
            dialog.CommonButtons = TaskDialogCommonButtons.Ok
            dialog.Show()
            return
        try:
            preselection = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()]
            elements_to_filter = self.all_elements if ('Name' in self.selected_filters or 'Comments' in self.selected_filters or 'Category' in self.selected_filters) else self.fab_elements
            if preselection:
                elements_to_filter = preselection
            filtered_ids = []
            for elem in elements_to_filter:
                if elem is None or not elem.IsValidObject:
                    continue
                elem_values = {prop: get_property_value(elem, prop, debug=False) for prop in self.selected_filters}
                matches = []
                for prop, filter_list in self.selected_filters.items():
                    elem_value = elem_values[prop]
                    prop_matches = []
                    for values, is_and in filter_list:
                        prop_matches.append((str(elem_value) in [str(v) for v in values], is_and))
                    and_matches_prop = [m for m, is_and in prop_matches if is_and]
                    or_matches_prop = [m for m, is_and in prop_matches if not is_and]
                    prop_result = (not and_matches_prop or all(and_matches_prop)) and \
                                 (not or_matches_prop or any(or_matches_prop))
                    matches.append(prop_result)
                if all(matches):
                    filtered_ids.append(elem.Id)
           
            if filtered_ids:
                element_id_list = List[DB.ElementId](filtered_ids)
                t = Transaction(doc, "Isolate Filtered Elements")
                t.Start()
                curview.IsolateElementsTemporary(element_id_list)
                t.Commit()
            else:
                dialog = TaskDialog("Warning")
                dialog.MainInstruction = "No elements match the selected filters."
                dialog.CommonButtons = TaskDialogCommonButtons.Ok
                dialog.Show()
        except Exception as e:
            dialog = TaskDialog("Error")
            dialog.MainInstruction = "Isolate Error: {}".format(str(e))
            dialog.CommonButtons = TaskDialogCommonButtons.Ok
            dialog.Show()
    def select_clicked(self, sender, args):
        if not self.selected_filters:
            dialog = TaskDialog("Warning")
            dialog.MainInstruction = "No filters selected to select."
            dialog.CommonButtons = TaskDialogCommonButtons.Ok
            dialog.Show()
            return
        try:
            preselection = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()]
            elements_to_filter = self.all_elements if ('Name' in self.selected_filters or 'Comments' in self.selected_filters or 'Category' in self.selected_filters) else self.fab_elements
            if preselection:
                elements_to_filter = preselection
            filtered_ids = []
            for elem in elements_to_filter:
                if elem is None or not elem.IsValidObject:
                    continue
                elem_values = {prop: get_property_value(elem, prop, debug=False) for prop in self.selected_filters}
                matches = []
                for prop, filter_list in self.selected_filters.items():
                    elem_value = elem_values[prop]
                    prop_matches = []
                    for values, is_and in filter_list:
                        prop_matches.append((str(elem_value) in [str(v) for v in values], is_and))
                    and_matches_prop = [m for m, is_and in prop_matches if is_and]
                    or_matches_prop = [m for m, is_and in prop_matches if not is_and]
                    prop_result = (not and_matches_prop or all(and_matches_prop)) and \
                                 (not or_matches_prop or any(or_matches_prop))
                    matches.append(prop_result)
                if all(matches):
                    filtered_ids.append(elem.Id)
           
            if filtered_ids:
                element_id_list = List[DB.ElementId](filtered_ids)
                uidoc.Selection.SetElementIds(element_id_list)
                self.Close()
            else:
                dialog = TaskDialog("Warning")
                dialog.MainInstruction = "No elements match the selected filters."
                dialog.CommonButtons = TaskDialogCommonButtons.Ok
                dialog.Show()
        except Exception as e:
            dialog = TaskDialog("Error")
            dialog.MainInstruction = "Select Error: {}".format(str(e))
            dialog.CommonButtons = TaskDialogCommonButtons.Ok
            dialog.Show()
    def cancel_clicked(self, sender, args):
        self.Close()
def get_property_value(elem, property_name, debug=False):
    if elem is None or not elem.IsValidObject:
        return None
    property_map = {
        'CID': lambda x: str(x.ItemCustomId) if x.ItemCustomId else None,
        'ServiceType': lambda x: Config.GetServiceTypeName(x.ServiceType) if x.ServiceType else None,
        'Name': lambda x: get_parameter_value_by_name_AsValueString(x, 'Family') or x.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString() if x.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM) else None,
        'Service Name': lambda x: get_parameter_value_by_name_AsString(x, 'Fabrication Service Name'),
        'Service Abbreviation': lambda x: get_parameter_value_by_name_AsString(x, 'Fabrication Service Abbreviation'),
        'Size': lambda x: get_parameter_value_by_name_AsString(x, 'Size of Primary End'),
        'STRATUS Assembly': lambda x: get_parameter_value_by_name_AsString(x, 'STRATUS Assembly'),
        'Line Number': lambda x: get_parameter_value_by_name_AsString(x, 'FP_Line Number'),
        'STRATUS Status': lambda x: get_parameter_value_by_name_AsString(x, 'STRATUS Status'),
        'Reference Level': lambda x: get_parameter_value_by_name_AsValueString(x, 'Reference Level'),
        'Item Number': lambda x: get_parameter_value_by_name_AsString(x, 'Item Number'),
        'Bundle Number': lambda x: get_parameter_value_by_name_AsString(x, 'FP_Bundle'),
        'REF BS Designation': lambda x: get_parameter_value_by_name_AsString(x, 'FP_REF BS Designation'),
        'REF Line Number': lambda x: get_parameter_value_by_name_AsString(x, 'FP_REF Line Number'),
        'Comments': lambda x: get_parameter_value_by_name_AsString(x, 'Comments'),
        'Specification': lambda x: Config.GetSpecificationName(x.Specification) if x.Specification else None,
        'Hanger Rod Size': lambda x: get_parameter_value_by_name_AsValueString(x, 'FP_Rod Size'),
        'Valve Number': lambda x: get_parameter_value_by_name_AsString(x, 'FP_Valve Number'),
        'Beam Hanger': lambda x: get_parameter_value_by_name_AsString(x, 'FP_Beam Hanger'),
        'Product Entry': lambda x: get_parameter_value_by_name_AsString(x, 'Product Entry'),
        'Category': lambda x: x.Category.Name if x.Category else None,
    }
    try:
        value = property_map.get(property_name, lambda x: None)(elem)
        return value
    except:
        return None
def get_parameter_id(property_name):
    param_map = {
        'STRATUS Assembly': 'STRATUS Assembly',
        'Line Number': 'FP_Line Number',
        'Service Name': 'Fabrication Service Name',
        'Service Abbreviation': 'Fabrication Service Abbreviation',
        'Size': 'Size of Primary End',
        'STRATUS Status': 'STRATUS Status',
        'Reference Level': 'Reference Level',
        'Item Number': 'Item Number',
        'Bundle Number': 'FP_Bundle',
        'REF BS Designation': 'FP_REF BS Designation',
        'REF Line Number': 'FP_REF Line Number',
        'Comments': 'Comments',
        'Hanger Rod Size': 'FP_Rod Size',
        'Valve Number': 'FP_Valve Number',
        'Beam Hanger': 'FP_Beam Hanger',
        'Product Entry': 'Product Entry',
        'Name': 'Family',
    }
    return param_map.get(property_name)
# Collect elements
preselection = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()]
fab_elements = preselection if preselection else FilteredElementCollector(doc, curview.Id) \
    .OfClass(DB.FabricationPart) \
    .WhereElementIsNotElementType() \
    .ToElements()
all_elements = preselection if preselection else FilteredElementCollector(doc, curview.Id) \
    .WhereElementIsNotElementType() \
    .ToElements()

# Build property options for fabrication parts
property_options = {}
for prop in ['CID', 'ServiceType', 'Service Name', 'Service Abbreviation', 'Size',
             'STRATUS Assembly', 'Line Number', 'STRATUS Status', 'Reference Level',
             'Item Number', 'Bundle Number', 'REF BS Designation', 'REF Line Number',
             'Specification', 'Hanger Rod Size', 'Valve Number', 'Beam Hanger', 'Product Entry']:
    values = set(filter(None, [get_property_value(elem, prop, debug=False) for elem in fab_elements]))
    if values:
        property_options[prop] = sorted(values)
# Build property options for all elements
for prop in ['Name', 'Comments', 'Category']:
    if all_elements:
        values = set(filter(None, [get_property_value(elem, prop, debug=False) for elem in all_elements]))
        if values:
            property_options[prop] = sorted(values)
if not property_options:
    dialog = TaskDialog("Error")
    dialog.MainInstruction = "No properties found for the selected elements."
    dialog.CommonButtons = TaskDialogCommonButtons.Ok
    dialog.Show()
    import sys
    sys.exit()
# Show form as modeless with DispatcherFrame
form = MultiPropertyFilterForm(property_options, fab_elements, all_elements)
form.frame = DispatcherFrame()
form.Closed += form.exit_frame
form.Show()
Dispatcher.PushFrame(form.frame)