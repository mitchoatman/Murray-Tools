import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('PresentationFramework')  # WPF assemblies
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
import System
from System.Windows import Window, Controls, Thickness, HorizontalAlignment, VerticalAlignment
from System.Windows.Controls import Grid, RowDefinition, ColumnDefinition, Label, TextBox, Button, ListBox, StackPanel
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, BuiltInCategory, FabricationPart, ElementId
from Autodesk.Revit.UI import UIApplication, TaskDialog
from Parameters.Get_Set_Params import get_parameter_value_by_name_AsString, set_parameter_by_name
from System.Collections.Generic import List
import re

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
app = UIApplication(doc.Application)

def natural_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', s)]

def get_fabrication_parts(view_id):
    categories = [BuiltInCategory.OST_FabricationPipework,
                  BuiltInCategory.OST_FabricationDuctwork,
                  BuiltInCategory.OST_FabricationContainment]
    parts = []
    for cat in categories:
        collector = FilteredElementCollector(doc, view_id).OfCategory(cat).OfClass(FabricationPart)
        parts.extend(list(collector))
    return parts

def get_assemblies():
    assemblies = set()
    for elem in get_fabrication_parts(doc.ActiveView.Id):
        param_asm = elem.LookupParameter("STRATUS Assembly")
        param_pkg = elem.LookupParameter("STRATUS Package")
        if param_asm and param_asm.HasValue:
            asm_val = param_asm.AsString()
            if asm_val:
                assemblies.add("Assembly: {}".format(asm_val))
        if param_pkg and param_pkg.HasValue:
            pkg_val = param_pkg.AsString()
            if pkg_val:
                assemblies.add("Package: {}".format(pkg_val))
    return sorted(assemblies, key=natural_key)

class RenameForm(Window):
    def __init__(self, assemblies):
        self.Title = "Rename STRATUS Assembly or Package"
        self.all_assemblies = assemblies
        self.Width = 300
        self.ResizeMode = System.Windows.ResizeMode.NoResize
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.selected_items = []
        self.InitializeComponents(assemblies)

    def update_assemblies_list(self, filter_text=""):
        # Store current selections
        self.selected_items = list(self.listbox.SelectedItems)
        self.listbox.Items.Clear()
        filtered = [a for a in self.all_assemblies if filter_text in a]
        for item in filtered:
            self.listbox.Items.Add(item)
        # Restore selections if they still exist in the filtered list
        for item in self.selected_items:
            if item in filtered:
                self.listbox.SelectedItems.Add(item)

    def on_search_text_changed(self, sender, event):
        self.update_assemblies_list(self.find_textbox.Text)

    def on_select_all_click(self, sender, event):
        self.listbox.SelectAll()

    def on_listbox_lost_focus(self, sender, event):
        # Preserve selections when focus is lost
        self.selected_items = list(self.listbox.SelectedItems)

    def on_rename_click(self, sender, event):
        find = self.find_textbox.Text
        replace = self.replace_textbox2.Text
        selected = list(self.listbox.SelectedItems)

        if not find or not replace:
            TaskDialog.Show("Warning", "Please enter both Find and Replace values.")
            return
        if not selected:
            TaskDialog.Show("Warning", "No assemblies selected.")
            return

        elements_to_process = []
        for elem in get_fabrication_parts(doc.ActiveView.Id):
            asm_val = get_parameter_value_by_name_AsString(elem, "STRATUS Assembly")
            pkg_val = get_parameter_value_by_name_AsString(elem, "STRATUS Package")
            for sel in selected:
                if sel.startswith("Assembly: ") and sel[10:] == asm_val:
                    elements_to_process.append((elem, "STRATUS Assembly"))
                elif sel.startswith("Package: ") and sel[9:] == pkg_val:
                    elements_to_process.append((elem, "STRATUS Package"))

        if not elements_to_process:
            TaskDialog.Show("Warning", "No fabrication parts matched your selection.")
            return

        t = Transaction(doc, "Rename STRATUS Parameters")
        t.Start()
        try:
            for elem, pname in elements_to_process:
                current = get_parameter_value_by_name_AsString(elem, pname)
                if current and not elem.LookupParameter(pname).IsReadOnly:
                    new_val = current.replace(find, replace)
                    set_parameter_by_name(elem, pname, new_val)
            t.Commit()
        except Exception as e:
            TaskDialog.Show("Error", "Error during transaction:\n{}".format(str(e)))
            if t.HasStarted():
                t.RollBack()
            return

        # Refresh assemblies with updated names and keep prefix
        self.all_assemblies = get_assemblies()
        self.update_assemblies_list(self.find_textbox.Text)

    def InitializeComponents(self, assemblies):
        # Create Grid layout
        grid = Grid()
        self.Content = grid

        # Define rows
        row_definitions = [
            RowDefinition(Height=System.Windows.GridLength.Auto),  # Find label
            RowDefinition(Height=System.Windows.GridLength.Auto),  # Find textbox + clear button
            RowDefinition(Height=System.Windows.GridLength.Auto),  # Replace label
            RowDefinition(Height=System.Windows.GridLength.Auto),  # Replace textbox
            RowDefinition(Height=System.Windows.GridLength.Auto),  # Listbox label
            RowDefinition(Height=System.Windows.GridLength(1, System.Windows.GridUnitType.Star)),  # Listbox
            RowDefinition(Height=System.Windows.GridLength.Auto)   # Buttons
        ]
        for row in row_definitions:
            grid.RowDefinitions.Add(row)

        # Define columns
        column_definitions = [
            ColumnDefinition(Width=System.Windows.GridLength(1, System.Windows.GridUnitType.Star)),  # Main content
            ColumnDefinition(Width=System.Windows.GridLength.Auto)  # Clear button
        ]
        grid.ColumnDefinitions.Add(column_definitions[0])
        grid.ColumnDefinitions.Add(column_definitions[1])

        # Calculate listbox height based on assemblies
        item_height = 20
        listbox_height = item_height * min(15, max(7, len(assemblies))) + 5
        self.Height = listbox_height + 250  # Increased by 10 for more spacing

        # Add controls
        row_index = 0

        # Find label
        find_label = Label()
        find_label.Content = "Find (Case Sensitive):"
        find_label.Margin = Thickness(10, 5, 10, 5)
        Grid.SetRow(find_label, row_index)
        Grid.SetColumnSpan(find_label, 2)
        grid.Children.Add(find_label)
        row_index += 1

        # Find textbox and clear button
        self.find_textbox = TextBox()
        self.find_textbox.Margin = Thickness(10, 0, 5, 5)
        self.find_textbox.TextChanged += self.on_search_text_changed
        Grid.SetRow(self.find_textbox, row_index)
        Grid.SetColumn(self.find_textbox, 0)
        grid.Children.Add(self.find_textbox)

        clear_button = Button()
        clear_button.Content = "X"
        clear_button.Width = 20
        clear_button.Height = 20
        clear_button.Margin = Thickness(0, 0, 10, 5)
        clear_button.Click += lambda sender, args: setattr(self.find_textbox, 'Text', '')
        Grid.SetRow(clear_button, row_index)
        Grid.SetColumn(clear_button, 1)
        grid.Children.Add(clear_button)
        row_index += 1

        # Replace label
        replace_label = Label()
        replace_label.Content = "Replace:"
        replace_label.Margin = Thickness(10, 0, 10, 5)
        Grid.SetRow(replace_label, row_index)
        Grid.SetColumnSpan(replace_label, 2)
        grid.Children.Add(replace_label)
        row_index += 1

        # Replace textbox
        self.replace_textbox2 = TextBox()
        self.replace_textbox2.Margin = Thickness(10, 0, 10, 5)
        Grid.SetRow(self.replace_textbox2, row_index)
        Grid.SetColumnSpan(self.replace_textbox2, 2)
        grid.Children.Add(self.replace_textbox2)
        row_index += 1

        # Listbox label
        listbox_label = Label()
        listbox_label.Content = "Select STRATUS Assemblies to Rename:"
        listbox_label.Margin = Thickness(10, 0, 10, 5)
        Grid.SetRow(listbox_label, row_index)
        Grid.SetColumnSpan(listbox_label, 2)
        grid.Children.Add(listbox_label)
        row_index += 1

        # Listbox
        self.listbox = ListBox()
        self.listbox.SelectionMode = System.Windows.Controls.SelectionMode.Extended
        self.listbox.Height = listbox_height
        self.listbox.Margin = Thickness(10, 0, 10, 0)
        for a in assemblies:
            self.listbox.Items.Add(a)
        self.listbox.LostFocus += self.on_listbox_lost_focus
        Grid.SetRow(self.listbox, row_index)
        Grid.SetColumnSpan(self.listbox, 2)
        grid.Children.Add(self.listbox)
        row_index += 1

        # Buttons
        button_panel = StackPanel()
        button_panel.Orientation = Controls.Orientation.Horizontal
        button_panel.HorizontalAlignment = HorizontalAlignment.Center
        button_panel.Margin = Thickness(0, 15, 0, 10)
        Grid.SetRow(button_panel, row_index)
        Grid.SetColumnSpan(button_panel, 2)
        grid.Children.Add(button_panel)

        self.select_all_button = Button()
        self.select_all_button.Content = "Select All"
        self.select_all_button.Width = 75
        self.select_all_button.Height = 25
        self.select_all_button.Margin = Thickness(5, 0, 5, 0)
        self.select_all_button.Click += self.on_select_all_click
        button_panel.Children.Add(self.select_all_button)

        self.rename_button = Button()
        self.rename_button.Content = "Rename"
        self.rename_button.Width = 75
        self.rename_button.Height = 25
        self.rename_button.Margin = Thickness(5, 0, 5, 0)
        self.rename_button.Click += self.on_rename_click
        button_panel.Children.Add(self.rename_button)

        self.cancel_button = Button()
        self.cancel_button.Content = "Close"
        self.cancel_button.Width = 75
        self.cancel_button.Height = 25
        self.cancel_button.Margin = Thickness(5, 0, 5, 0)
        self.cancel_button.Click += lambda sender, args: self.Close()
        button_panel.Children.Add(self.cancel_button)

# Run form
form = RenameForm(get_assemblies())
form.ShowDialog()