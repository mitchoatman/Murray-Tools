# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, FabricationPart
from Autodesk.Revit.UI import TaskDialog, TaskDialogResult, TaskDialogCommonButtons
import math

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

# WPF imports (proven pattern)
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

import System
from System.Windows import (
    Window, Thickness, HorizontalAlignment, VerticalAlignment,
    WindowStartupLocation, ResizeMode
)
from System.Windows.Controls import (
    Label, TextBox, Button, Grid,
    RowDefinition, ColumnDefinition
)
from System.Windows import GridLength

class FuzzDistanceWindow(Window):
    def __init__(self):
        self.Title = "Fuzz Distance"
        self.Width = 260
        self.Height = 160
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.NoResize
        self.Topmost = True

        # Main Grid
        grid = Grid()
        grid.Margin = Thickness(10)

        # Rows
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(40)))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(50)))

        # Columns
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(80)))
        grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, System.Windows.GridUnitType.Star)))

        # Label
        lbl = Label()
        lbl.Content = "Fuzz:"
        lbl.FontSize = 14
        lbl.VerticalAlignment = VerticalAlignment.Center
        lbl.Margin = Thickness(25, 0, 0, 0)
        Grid.SetRow(lbl, 0)
        Grid.SetColumn(lbl, 0)
        grid.Children.Add(lbl)

        # TextBox
        self.txtFuzz = TextBox()
        self.txtFuzz.Text = "0.0625"
        self.txtFuzz.FontSize = 14
        self.txtFuzz.Width = 120
        self.txtFuzz.HorizontalAlignment = HorizontalAlignment.Left
        self.txtFuzz.VerticalAlignment = VerticalAlignment.Center
        self.txtFuzz.Margin = Thickness(1, 0, 0, 0)
        Grid.SetRow(self.txtFuzz, 0)
        Grid.SetColumn(self.txtFuzz, 1)
        grid.Children.Add(self.txtFuzz)

        # Button
        btn = Button()
        btn.Content = "Set Distance"
        btn.Width = 100
        btn.Height = 30
        btn.HorizontalAlignment = HorizontalAlignment.Center
        btn.IsDefault = True
        btn.Click += self.on_ok_click
        Grid.SetRow(btn, 1)
        Grid.SetColumn(btn, 0)
        Grid.SetColumnSpan(btn, 2)
        grid.Children.Add(btn)

        self.Content = grid
        self.result = None

    def on_ok_click(self, sender, args):
        try:
            value = float(self.txtFuzz.Text)
            if value <= 0:
                raise ValueError
            self.result = value / 12.0          # inches → feet
            self.DialogResult = True
        except ValueError:
            TaskDialog.Show("Invalid Input", "Please enter a valid positive number in decimal inches.")
            self.DialogResult = False

# Show dialog
window = FuzzDistanceWindow()
if not window.ShowDialog() or window.result is None:
    TaskDialog.Show("Cancelled", "Operation cancelled or invalid input.")
    import sys
    sys.exit()

fuzz_distance = window.result

# ---------- Core logic ----------
def GetCenterPoint(element_id):
    elem = doc.GetElement(element_id)
    bbox = elem.get_BoundingBox(None)
    if bbox and bbox.Enabled:
        center = (bbox.Max + bbox.Min) / 2
        return (center.X, center.Y, center.Z)
    return None

def calculate_distance(p1, p2):
    if p1 is None or p2 is None:
        return float('inf')
    return math.sqrt((p1[0] - p2[0])**2 + (p1[1] - p2[1])**2 + (p1[2] - p2[2])**2)

# Collect visible Fabrication Parts (fixed syntax)
fabrication_parts = FilteredElementCollector(doc, curview.Id) \
                    .OfClass(FabricationPart) \
                    .WhereElementIsNotElementType() \
                    .ToElements()

center_points = []
element_ids = []

for part in fabrication_parts:
    cp = GetCenterPoint(part.Id)
    if cp:
        center_points.append(cp)
        element_ids.append(part.Id)

# Identify duplicates
duplicate_ids = []
unique_centers = []

for i, cp in enumerate(center_points):
    if any(calculate_distance(cp, uc) <= fuzz_distance for uc in unique_centers):
        duplicate_ids.append(element_ids[i])
    else:
        unique_centers.append(cp)

# Delete duplicates with confirmation using TaskDialog
if duplicate_ids:
    td = TaskDialog("Confirm Deletion")
    td.MainInstruction = "Delete {} duplicate fabrication part(s)?".format(len(duplicate_ids))
    td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
    td.DefaultButton = TaskDialogResult.Yes

    if td.Show() == TaskDialogResult.Yes:
        with Transaction(doc, "Delete Duplicate Fabrication Parts") as t:
            t.Start()
            for eid in duplicate_ids:
                try:
                    doc.Delete(eid)
                except:
                    pass
            t.Commit()
        TaskDialog.Show("Success", "Duplicate(s) successfully removed.")
    else:
        TaskDialog.Show("Cancelled", "Operation cancelled by user.")
else:
    TaskDialog.Show("No Duplicates", "No duplicate fabrication parts found within the specified fuzz distance.")