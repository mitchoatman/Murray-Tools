import os
# Folder and file for storing seismic brace defaults
folder_name = "C:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_SeismicAutoBrace.txt')

# Ensure folder exists
if not os.path.exists(folder_name):
    os.makedirs(folder_name)

# Default spacing values (feet)
default_transverse = 20
default_longitudinal = 40

# Create file if missing and write default values
if not os.path.exists(filepath):
    with open(filepath, 'w') as f:
        f.write("{},{}".format(default_transverse, default_longitudinal))

# Read values from file
try:
    with open(filepath, 'r') as f:
        data = f.read().split(',')
        transverse_spacing = float(data[0])
        longitudinal_spacing = float(data[1])
except (ValueError, IndexError, FileNotFoundError):
    transverse_spacing = default_transverse
    longitudinal_spacing = default_longitudinal

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import Window, Thickness, WindowStyle, ResizeMode, WindowStartupLocation, HorizontalAlignment
from System.Windows.Controls import Label, TextBox, Button, Grid, RowDefinition

class SpacingForm(Window):
    def __init__(self, transverse_spacing, longitudinal_spacing):
        self.Title = "Seismic Brace Spacing"
        self.Width = 300
        self.Height = 200
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.result = None

        grid = Grid()
        grid.Margin = Thickness(10)
        for _ in range(5):
            grid.RowDefinitions.Add(RowDefinition())

        self.Content = grid

        label1 = Label()
        label1.Content = "Transverse Spacing:"
        Grid.SetRow(label1, 0)
        grid.Children.Add(label1)

        self.transverse_textbox = TextBox()
        self.transverse_textbox.Text = str(transverse_spacing)  # Set initial value
        self.transverse_textbox.Margin = Thickness(0, 0, 0, 5)
        Grid.SetRow(self.transverse_textbox, 1)
        grid.Children.Add(self.transverse_textbox)

        label2 = Label()
        label2.Content = "Longitudinal Spacing:"
        Grid.SetRow(label2, 2)
        grid.Children.Add(label2)

        self.longitudinal_textbox = TextBox()
        self.longitudinal_textbox.Text = str(longitudinal_spacing)  # Set initial value
        self.longitudinal_textbox.Margin = Thickness(0, 0, 0, 5)
        Grid.SetRow(self.longitudinal_textbox, 3)
        grid.Children.Add(self.longitudinal_textbox)

        ok_button = Button()
        ok_button.Content = "OK"
        ok_button.Width = 75
        ok_button.Height = 25
        ok_button.HorizontalAlignment = HorizontalAlignment.Center
        ok_button.Click += self.on_ok
        Grid.SetRow(ok_button, 4)
        grid.Children.Add(ok_button)

        self.Closing += self.on_closing

    def on_ok(self, sender, args):
        self.result = {
            "transverse_spacing": self.transverse_textbox.Text,
            "longitudinal_spacing": self.longitudinal_textbox.Text
        }
        self.DialogResult = True
        self.Close()

    def on_closing(self, sender, args):
        if not self.DialogResult:
            self.result = None

# Show dialog
form = SpacingForm(transverse_spacing, longitudinal_spacing)  # note: init no longer takes args
form.ShowDialog()

if not form.DialogResult or form.result is None:
    pass
else:
    # Parse user input (supports feet-inches formats like the reference script)
    def parse_spacing(input_str):
        try:
            input_str = input_str.strip()
            if '-' in input_str:           # 20-6
                feet, inches = input_str.split('-')
                feet = float(feet.strip("'"))
                inches = float(inches.strip('"')) / 12.0
                return feet + inches
            elif "'" in input_str or '"' in input_str:  # 20'  20'-6"
                input_str = input_str.replace("'", "").replace('"', "")
                parts = input_str.split('-')
                feet = float(parts[0])
                inches = float(parts[1]) / 12.0 if len(parts) > 1 else 0.0
                return feet + inches
            else:                          # 20.5   20
                return float(input_str)
        except (ValueError, IndexError):
            return None

    transverse_input = form.result["transverse_spacing"]
    longitudinal_input = form.result["longitudinal_spacing"]

    transverse_spacing = parse_spacing(transverse_input)
    longitudinal_spacing = parse_spacing(longitudinal_input)

    # Basic validation
    if transverse_spacing is None or longitudinal_spacing is None or \
       transverse_spacing <= 0 or longitudinal_spacing <= 0:
        # You can replace with TaskDialog if running in Revit
        print("Invalid spacing values. Please enter positive numbers.")
        # Or keep previous values:
        # transverse_spacing = default_transverse
        # longitudinal_spacing = default_longitudinal
    else:
        # Save valid values back to file (as decimal feet)
        with open(filepath, 'w') as f:
            f.write("{},{}".format(transverse_spacing, longitudinal_spacing))