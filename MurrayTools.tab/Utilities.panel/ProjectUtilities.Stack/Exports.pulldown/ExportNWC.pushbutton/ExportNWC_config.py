import os
import json
import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
clr.AddReference("System")

from System.Windows.Forms import *
from System.Drawing import Point, Size, Font, FontStyle
from System import Array

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
    "ConvertLinkedCADFormats": False
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
class ExportOptionsForm(Form):
    def __init__(self):
        self.Text = "NWC Export Options"
        self.Width = 400  # Increased width to accommodate wider checkboxes
        self.Height = 360  # Increased height for title label and ComboBox
        self.FormBorderStyle = 0
        self.StartPosition = FormStartPosition.CenterScreen
        self.options = load_options()
        self.create_controls()

    def create_controls(self):
        y = 20
        spacing = 30
        checkbox_width = 300  # Set width to prevent text wrapping

        # Title label
        self.title_label = Label()
        self.title_label.Text = "NWC Export Options"
        self.title_label.Location = Point(20, y)
        self.title_label.Size = Size(300, 20)
        self.title_label.Font = Font("Arial", 12, FontStyle.Bold)  # Bold and larger for emphasis
        self.Controls.Add(self.title_label)
        y += spacing

        # ComboBox for Coordinates
        self.combobox_coordinates = ComboBox()
        self.combobox_coordinates.Items.AddRange(Array[str](["Shared", "Project Internal"]))
        self.combobox_coordinates.SelectedItem = self.options["Coordinates"]
        self.combobox_coordinates.Location = Point(20, y)
        self.combobox_coordinates.Size = Size(150, 20)
        self.combobox_coordinates.DropDownStyle = ComboBoxStyle.DropDownList  # Prevent typing
        self.Controls.Add(self.combobox_coordinates)
        y += spacing

        # Checkboxes for boolean options
        self.checkbox_find_materials = CheckBox()
        self.checkbox_find_materials.Text = "Find Missing Materials"
        self.checkbox_find_materials.Checked = self.options["FindMissingMaterials"]
        self.checkbox_find_materials.Location = Point(20, y)
        self.checkbox_find_materials.Size = Size(checkbox_width, 20)
        self.Controls.Add(self.checkbox_find_materials)
        y += spacing

        self.checkbox_element_properties = CheckBox()
        self.checkbox_element_properties.Text = "Convert Element Properties"
        self.checkbox_element_properties.Checked = self.options["ConvertElementProperties"]
        self.checkbox_element_properties.Location = Point(20, y)
        self.checkbox_element_properties.Size = Size(checkbox_width, 20)
        self.Controls.Add(self.checkbox_element_properties)
        y += spacing

        self.checkbox_export_urls = CheckBox()
        self.checkbox_export_urls.Text = "Export URLs"
        self.checkbox_export_urls.Checked = self.options["ExportUrls"]
        self.checkbox_export_urls.Location = Point(20, y)
        self.checkbox_export_urls.Size = Size(checkbox_width, 20)
        self.Controls.Add(self.checkbox_export_urls)
        y += spacing

        self.checkbox_linked_cad = CheckBox()
        self.checkbox_linked_cad.Text = "Convert Linked CAD Formats"
        self.checkbox_linked_cad.Checked = self.options["ConvertLinkedCADFormats"]
        self.checkbox_linked_cad.Location = Point(20, y)
        self.checkbox_linked_cad.Size = Size(checkbox_width, 20)
        self.Controls.Add(self.checkbox_linked_cad)
        y += spacing

        self.checkbox_divide_levels = CheckBox()
        self.checkbox_divide_levels.Text = "Divide File Into Levels"
        self.checkbox_divide_levels.Checked = self.options["DivideFileIntoLevels"]
        self.checkbox_divide_levels.Location = Point(20, y)
        self.checkbox_divide_levels.Size = Size(checkbox_width, 20)
        self.Controls.Add(self.checkbox_divide_levels)
        y += spacing

        # Save button
        self.save_button = Button()
        self.save_button.Text = "Save"
        self.save_button.Location = Point(20, y + 20)
        self.save_button.Size = Size(80, 25)  # Wider button
        self.save_button.Click += self.save_button_click
        self.Controls.Add(self.save_button)

        # Cancel button
        self.cancel_button = Button()
        self.cancel_button.Text = "Cancel"
        self.cancel_button.Location = Point(110, y + 20)  # Adjusted position
        self.cancel_button.Size = Size(80, 25)  # Wider button
        self.cancel_button.Click += self.cancel_button_click
        self.Controls.Add(self.cancel_button)

    def save_button_click(self, sender, args):
        self.options["Coordinates"] = self.combobox_coordinates.SelectedItem
        self.options["FindMissingMaterials"] = self.checkbox_find_materials.Checked
        self.options["DivideFileIntoLevels"] = self.checkbox_divide_levels.Checked
        self.options["ConvertElementProperties"] = self.checkbox_element_properties.Checked
        self.options["ExportUrls"] = self.checkbox_export_urls.Checked
        self.options["ConvertLinkedCADFormats"] = self.checkbox_linked_cad.Checked
        save_options(self.options)
        self.DialogResult = DialogResult.OK
        self.Close()

    def cancel_button_click(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()

# Run the dialog
def show_export_options_dialog():
    form = ExportOptionsForm()
    return form.ShowDialog()

if __name__ == "__main__":
    Application.EnableVisualStyles()
    result = show_export_options_dialog()