# -*- coding: UTF-8 -*-
import os
import json
import clr
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('System')
clr.AddReference('System.Windows.Forms')  # For DialogResult
from System.Windows import Application, Window, Thickness, HorizontalAlignment, VerticalAlignment, ResizeMode, WindowStartupLocation, GridLength, GridUnitType
from System.Windows.Controls import Button, ComboBox, CheckBox, Grid, RowDefinition, ColumnDefinition, Label, StackPanel, ScrollViewer, Orientation, ScrollBarVisibility
from System.Windows.Media import Brushes, FontFamily
from System.Windows.Controls.Primitives import UniformGrid
from System.Windows.Forms import DialogResult

# Path to save options
OPTIONS_FILE = r"C:\Temp\Ribbon_NavisworksExportOptions.txt"
FOLDER_PATH = r"C:\Temp"

# Ensure folder exists
if not os.path.exists(FOLDER_PATH):
    os.makedirs(FOLDER_PATH)

# Default options
default_options = {
    "ExportScope": "View",
    "Coordinates": "Shared",
    "FindMissingMaterials": True,
    "DivideFileIntoLevels": False,
    "ConvertElementProperties": True,
    "ExportUrls": False,
    "ConvertLinkedCADFormats": False,
    "ExportLinks": False,
    "SuccessMessage": True
}

# Load saved options if they exist
def load_options():
    if os.path.exists(OPTIONS_FILE):
        with open(OPTIONS_FILE, 'r') as f:
            return json.load(f)
    return default_options

# Save options to file
def save_options(options):
    with open(OPTIONS_FILE, 'w') as f:
        json.dump(options, f, indent=4)

# Create dialog form
class ExportOptionsForm(Window):
    def __init__(self):
        self.Title = "NWC Export Options"
        self.Width = 400
        self.Height = 300
        self.MinWidth = 400
        self.MinHeight = 300
        self.ResizeMode = ResizeMode.CanResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.options = load_options()
        self.InitializeComponents()

    def InitializeComponents(self):
        grid = Grid()
        grid.Margin = Thickness(10)
        grid.VerticalAlignment = VerticalAlignment.Stretch
        grid.HorizontalAlignment = HorizontalAlignment.Stretch

        for i in range(9):
            row = RowDefinition()
            row.Height = GridLength.Auto
            grid.RowDefinitions.Add(row)
        grid.ColumnDefinitions.Add(ColumnDefinition())

        y = 0
        spacing = 30

        # ComboBox for Coordinates
        label_coordinates = Label()
        label_coordinates.Content = "Coordinates:"
        label_coordinates.FontFamily = FontFamily("Arial")
        label_coordinates.FontSize = 12
        Grid.SetRow(label_coordinates, y)
        grid.Children.Add(label_coordinates)
        y += 1

        self.combobox_coordinates = ComboBox()
        self.combobox_coordinates.ItemsSource = ["Shared", "Project Internal"]
        self.combobox_coordinates.SelectedItem = self.options["Coordinates"]
        self.combobox_coordinates.FontFamily = FontFamily("Arial")
        self.combobox_coordinates.FontSize = 12
        self.combobox_coordinates.Margin = Thickness(0, 0, 0, 5)
        Grid.SetRow(self.combobox_coordinates, y)
        grid.Children.Add(self.combobox_coordinates)
        y += 1

        # Checkboxes for boolean options
        self.checkbox_find_materials = CheckBox()
        self.checkbox_find_materials.Content = "Find Missing Materials"
        self.checkbox_find_materials.IsChecked = self.options["FindMissingMaterials"]
        self.checkbox_find_materials.FontFamily = FontFamily("Arial")
        self.checkbox_find_materials.FontSize = 12
        self.checkbox_find_materials.Margin = Thickness(0, 0, 0, 5)
        Grid.SetRow(self.checkbox_find_materials, y)
        grid.Children.Add(self.checkbox_find_materials)
        y += 1

        self.checkbox_element_properties = CheckBox()
        self.checkbox_element_properties.Content = "Convert Element Properties"
        self.checkbox_element_properties.IsChecked = self.options["ConvertElementProperties"]
        self.checkbox_element_properties.FontFamily = FontFamily("Arial")
        self.checkbox_element_properties.FontSize = 12
        self.checkbox_element_properties.Margin = Thickness(0, 0, 0, 5)
        Grid.SetRow(self.checkbox_element_properties, y)
        grid.Children.Add(self.checkbox_element_properties)
        y += 1

        self.checkbox_export_urls = CheckBox()
        self.checkbox_export_urls.Content = "Export URLs"
        self.checkbox_export_urls.IsChecked = self.options["ExportUrls"]
        self.checkbox_export_urls.FontFamily = FontFamily("Arial")
        self.checkbox_export_urls.FontSize = 12
        self.checkbox_export_urls.Margin = Thickness(0, 0, 0, 5)
        Grid.SetRow(self.checkbox_export_urls, y)
        grid.Children.Add(self.checkbox_export_urls)
        y += 1

        self.checkbox_linked_cad = CheckBox()
        self.checkbox_linked_cad.Content = "Convert Linked CAD Formats"
        self.checkbox_linked_cad.IsChecked = self.options["ConvertLinkedCADFormats"]
        self.checkbox_linked_cad.FontFamily = FontFamily("Arial")
        self.checkbox_linked_cad.FontSize = 12
        self.checkbox_linked_cad.Margin = Thickness(0, 0, 0, 5)
        Grid.SetRow(self.checkbox_linked_cad, y)
        grid.Children.Add(self.checkbox_linked_cad)
        y += 1

        self.checkbox_export_links = CheckBox()
        self.checkbox_export_links.Content = "Convert Linked Files"
        self.checkbox_export_links.IsChecked = self.options.get("ExportLinks", False)
        self.checkbox_export_links.FontFamily = FontFamily("Arial")
        self.checkbox_export_links.FontSize = 12
        self.checkbox_export_links.Margin = Thickness(0, 0, 0, 5)
        Grid.SetRow(self.checkbox_export_links, y)
        grid.Children.Add(self.checkbox_export_links)
        y += 1

        self.checkbox_divide_levels = CheckBox()
        self.checkbox_divide_levels.Content = "Divide File Into Levels"
        self.checkbox_divide_levels.IsChecked = self.options["DivideFileIntoLevels"]
        self.checkbox_divide_levels.FontFamily = FontFamily("Arial")
        self.checkbox_divide_levels.FontSize = 12
        self.checkbox_divide_levels.Margin = Thickness(0, 0, 0, 5)
        Grid.SetRow(self.checkbox_divide_levels, y)
        grid.Children.Add(self.checkbox_divide_levels)
        y += 1

        self.checkbox_success_message = CheckBox()
        self.checkbox_success_message.Content = "Show Success Message"
        self.checkbox_success_message.IsChecked = self.options.get("SuccessMessage", True)
        self.checkbox_success_message.FontFamily = FontFamily("Arial")
        self.checkbox_success_message.FontSize = 12
        self.checkbox_success_message.Margin = Thickness(0, 0, 0, 5)
        Grid.SetRow(self.checkbox_success_message, y)
        grid.Children.Add(self.checkbox_success_message)
        y += 1

        # Button panel
        button_panel = UniformGrid()
        button_panel.Columns = 2
        button_panel.HorizontalAlignment = HorizontalAlignment.Center
        button_panel.Margin = Thickness(0, 30, 0, 0)

        self.save_button = Button()
        self.save_button.Content = "Save"
        self.save_button.Width = 80
        self.save_button.FontFamily = FontFamily("Arial")
        self.save_button.FontSize = 12
        self.save_button.Margin = Thickness(0, 0, 10, 0)
        self.save_button.Click += self.save_button_click
        button_panel.Children.Add(self.save_button)

        self.cancel_button = Button()
        self.cancel_button.Content = "Cancel"
        self.cancel_button.Width = 80
        self.cancel_button.FontFamily = FontFamily("Arial")
        self.cancel_button.FontSize = 12
        self.cancel_button.Margin = Thickness(10, 0, 0, 0)
        self.cancel_button.Click += self.cancel_button_click
        button_panel.Children.Add(self.cancel_button)

        Grid.SetRow(button_panel, y)
        grid.Children.Add(button_panel)

        self.Content = grid

    def save_button_click(self, sender, args):
        self.options["Coordinates"] = self.combobox_coordinates.SelectedItem
        self.options["FindMissingMaterials"] = self.checkbox_find_materials.IsChecked
        self.options["DivideFileIntoLevels"] = self.checkbox_divide_levels.IsChecked
        self.options["ConvertElementProperties"] = self.checkbox_element_properties.IsChecked
        self.options["ExportUrls"] = self.checkbox_export_urls.IsChecked
        self.options["ConvertLinkedCADFormats"] = self.checkbox_linked_cad.IsChecked
        self.options["ExportLinks"] = self.checkbox_export_links.IsChecked
        self.options["SuccessMessage"] = self.checkbox_success_message.IsChecked
        save_options(self.options)
        self.DialogResult = True
        self.Close()

    def cancel_button_click(self, sender, args):
        self.DialogResult = False
        self.Close()

# Run the dialog
def show_export_options_dialog():
    form = ExportOptionsForm()
    form.ShowDialog()
    return DialogResult.OK if form.DialogResult else DialogResult.Cancel

if __name__ == "__main__":
    result = show_export_options_dialog()