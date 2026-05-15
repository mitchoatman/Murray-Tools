# -*- coding: utf-8 -*-

import os
import sys
import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from Autodesk.Revit import DB
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

from System.Windows import (
    Window,
    Thickness,
    WindowStyle,
    ResizeMode,
    WindowStartupLocation,
    HorizontalAlignment
)

from System.Windows.Controls import (
    Label,
    TextBox,
    Button,
    Grid,
    RowDefinition
)

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


# -----------------------------
# Config File
# -----------------------------
folder_name = r"C:\Temp"
filepath = os.path.join(folder_name, "Ribbon_SetBearerExtn.txt")

if not os.path.exists(folder_name):
    os.makedirs(folder_name)

# Default = 2 inches
if not os.path.exists(filepath):
    with open(filepath, "w") as f:
        f.write("2")


# -----------------------------
# Selection Filter
# -----------------------------
class FabricationHangerSelectionFilter(ISelectionFilter):

    def AllowElement(self, element):

        cat = element.Category

        return (
            cat
            and cat.Name == "MEP Fabrication Hangers"
        )

    def AllowReference(self, reference, point):
        return False


# -----------------------------
# WPF Dialog
# -----------------------------
class BearerExtensionForm(Window):

    def __init__(self, initial_value):

        self.Title = "Modify Bearer Extension"
        self.Width = 320
        self.Height = 165
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen

        self.result_value = None

        grid = Grid()
        grid.Margin = Thickness(10)

        for _ in range(3):
            grid.RowDefinitions.Add(RowDefinition())

        self.Content = grid

        # Label
        label = Label()
        label.Content = "New Bearer Extension (Inches):"
        label.Margin = Thickness(0, 0, 0, 5)

        Grid.SetRow(label, 0)
        grid.Children.Add(label)

        # Textbox
        self.textbox = TextBox()
        self.textbox.Text = str(round(initial_value, 3))
        self.textbox.Margin = Thickness(0, 0, 0, 10)

        Grid.SetRow(self.textbox, 1)
        grid.Children.Add(self.textbox)

        # OK Button
        ok_button = Button()
        ok_button.Content = "OK"
        ok_button.Width = 75
        ok_button.Height = 25
        ok_button.HorizontalAlignment = HorizontalAlignment.Center
        ok_button.Click += self.ok_clicked

        Grid.SetRow(ok_button, 2)
        grid.Children.Add(ok_button)

        self.textbox.Focus()
        self.textbox.SelectAll()

    def ok_clicked(self, sender, args):

        try:

            self.result_value = float(self.textbox.Text)

            self.DialogResult = True
            self.Close()

        except:

            TaskDialog.Show(
                "Invalid Input",
                "Please enter a valid numeric value."
            )


# -----------------------------
# Select Hangers
# -----------------------------
try:

    selected_refs = uidoc.Selection.PickObjects(
        ObjectType.Element,
        FabricationHangerSelectionFilter(),
        "Select Fabrication Hangers"
    )

except:
    sys.exit()

hangers = [
    doc.GetElement(ref.ElementId)
    for ref in selected_refs
]

if not hangers:

    TaskDialog.Show(
        "Modify Bearer Extension",
        "No fabrication hangers were selected."
    )

    sys.exit()


# -----------------------------
# Read Previous Value
# -----------------------------
with open(filepath, "r") as f:

    previous_value = float(f.read().strip())


# -----------------------------
# Show Dialog
# -----------------------------
form = BearerExtensionForm(previous_value)

if not form.ShowDialog():
    sys.exit()


# -----------------------------
# Get New Value
# -----------------------------
new_extension_inches = form.result_value
new_extension_feet = new_extension_inches / 12.0


# -----------------------------
# Save Value
# -----------------------------
with open(filepath, "w") as f:

    f.write(str(new_extension_inches))


# -----------------------------
# Apply Changes
# -----------------------------
t = DB.Transaction(
    doc,
    "Modify Bearer Extension"
)

t.Start()

failed = []

for hanger in hangers:

    try:

        rod_info = hanger.GetRodInfo()
        rod_count = rod_info.RodCount

        for rod_index in range(rod_count):

            rod_info.SetBearerExtension(
                rod_index,
                new_extension_feet
            )

    except:

        failed.append(str(hanger.Id))

t.Commit()


# -----------------------------
# Report Failures
# -----------------------------
if failed:

    TaskDialog.Show(
        "Modify Bearer Extension",
        "Failed to update {} hanger(s).".format(
            len(failed)
        )
    )