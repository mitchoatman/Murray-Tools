# -*- coding: utf-8 -*-
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('RevitAPIUI')

import System
from Autodesk.Revit.DB import BoundingBoxXYZ, FilteredElementCollector, Transaction, BuiltInCategory, FabricationPart, LogicalOrFilter, ElementCategoryFilter
from Autodesk.Revit.UI import TaskDialog
from System import Array
import math

from System.Windows import Window, WindowStartupLocation
from System.Windows.Controls import Label, ComboBox, Button, TextBlock, Grid, StackPanel
from System.Windows import Thickness, HorizontalAlignment, VerticalAlignment

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView

# Collect levels
level_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType()
level_elevations = {"(None)": None}
level_ids = {}
for level in level_collector:
    level_elevations[level.Name] = level.Elevation
    level_ids[level.Name] = level.Id

sorted_levels = sorted(level_elevations.items(), key=lambda x: x[1] if x[1] is not None else float('inf'))

if len(sorted_levels) > 1:
    pre_set_bottom_level = sorted_levels[0][0]
    pre_set_top_level = sorted_levels[1][0]
else:
    pre_set_bottom_level = sorted_levels[0][0] if sorted_levels else "(None)"
    pre_set_top_level = "(None)"

def get_bottom_point_z(ele):
    bBox = doc.GetElement(ele).get_BoundingBox(None)
    if bBox is None:
        return None
    return bBox.Min.Z

# WPF Grid Dialog with OK/Cancel
class LevelSelectionWindow(Window):
    def __init__(self, levels, default_top, default_bottom):
        self.Title = "Select Levels"
        self.Width = 400
        self.Height = 270
        self.ResizeMode = System.Windows.ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen

        grid = Grid()
        grid.Margin = Thickness(15)

        # Define rows
        for i in range(8):
            grid.RowDefinitions.Add(System.Windows.Controls.RowDefinition())
        grid.ColumnDefinitions.Add(System.Windows.Controls.ColumnDefinition())
        grid.ColumnDefinitions.Add(System.Windows.Controls.ColumnDefinition())

        # Top Level
        lbl_top = Label(Content="Choose Top Level:")
        Grid.SetRow(lbl_top, 0); Grid.SetColumn(lbl_top, 0)
        grid.Children.Add(lbl_top)

        self.top_combo = ComboBox()
        for l in sorted(levels.keys()):
            self.top_combo.Items.Add(l)
        self.top_combo.SelectedItem = default_top
        self.top_combo.ToolTip = "Set this to (None) for Roof level"
        Grid.SetRow(self.top_combo, 0); Grid.SetColumn(self.top_combo, 1)
        grid.Children.Add(self.top_combo)

        # Bottom Level
        lbl_bot = Label(Content="Choose Bottom Level:")
        Grid.SetRow(lbl_bot, 1); Grid.SetColumn(lbl_bot, 0)
        grid.Children.Add(lbl_bot)

        self.bottom_combo = ComboBox()
        for l in sorted(levels.keys()):
            self.bottom_combo.Items.Add(l)
        self.bottom_combo.SelectedItem = default_bottom
        self.bottom_combo.ToolTip = "Set this to (None) for Underground level"
        Grid.SetRow(self.bottom_combo, 1); Grid.SetColumn(self.bottom_combo, 1)
        grid.Children.Add(self.bottom_combo)

        # Instructions
        inst1 = TextBlock(Text="Parts are assigned to Bottom when both levels are set.")
        inst2 = TextBlock(Text="Parts are assigned to Bottom if Top is (None) ROOF.")
        inst3 = TextBlock(Text="Parts are assigned to Top if Bottom is (None) UG.")

        Grid.SetRow(inst1, 2); Grid.SetColumnSpan(inst1, 2)
        # Grid.SetRow(inst2, 3); Grid.SetColumnSpan(inst2, 2)
        # Grid.SetRow(inst3, 4); Grid.SetColumnSpan(inst3, 2)
        grid.Children.Add(inst1) #; grid.Children.Add(inst2); grid.Children.Add(inst3)

        # Buttons Panel
        btn_panel = StackPanel(Orientation=System.Windows.Controls.Orientation.Horizontal,
                               HorizontalAlignment=HorizontalAlignment.Center,
                               Margin=Thickness(0, 0, 0, 0))

        self.ok_button = Button(Content="Ok", Width=80, Height=25, Margin=Thickness(5,0,5,0))
        self.ok_button.Click += self.on_ok
        btn_panel.Children.Add(self.ok_button)

        self.cancel_button = Button(Content="Cancel", Width=80, Height=25, Margin=Thickness(5,0,5,0))
        self.cancel_button.Click += self.on_cancel
        btn_panel.Children.Add(self.cancel_button)

        Grid.SetRow(btn_panel, 6); Grid.SetColumnSpan(btn_panel, 2)
        grid.Children.Add(btn_panel)

        self.Content = grid
        self.result = None

    def on_ok(self, sender, event):
        self.result = {
            "TopLevel": self.top_combo.SelectedItem,
            "BotLevel": self.bottom_combo.SelectedItem
        }
        self.DialogResult = True
        self.Close()

    def on_cancel(self, sender, event):
        self.result = None
        self.DialogResult = False
        self.Close()

try:
    form = LevelSelectionWindow(level_elevations, pre_set_top_level, pre_set_bottom_level)
    if form.ShowDialog() and form.result:
        TopLevelName = form.result["TopLevel"]
        BotLevelName = form.result["BotLevel"]

        TopLevelElev = level_elevations[TopLevelName]
        BotLevelElev = level_elevations[BotLevelName]
        TopLevelId = level_ids.get(TopLevelName, None)
        BotLevelId = level_ids.get(BotLevelName, None)

        pipework_filter = ElementCategoryFilter(BuiltInCategory.OST_FabricationPipework)
        ductwork_filter = ElementCategoryFilter(BuiltInCategory.OST_FabricationDuctwork)
        combined_filter = LogicalOrFilter(pipework_filter, ductwork_filter)

        fabrication_elements = FilteredElementCollector(doc, curview.Id) \
                               .OfClass(FabricationPart) \
                               .WherePasses(combined_filter) \
                               .WhereElementIsNotElementType() \
                               .ToElements()

        t = Transaction(doc, "Assign Levels to Fabrication Parts")
        t.Start()
        for elem in fabrication_elements:
            try:
                bottom_z = get_bottom_point_z(elem.Id)
                if bottom_z is None:
                    continue

                if BotLevelElev is not None and TopLevelElev is not None:
                    if BotLevelElev <= bottom_z <= TopLevelElev:
                        elem.LookupParameter("Reference Level").Set(BotLevelId)

                elif TopLevelElev is None and BotLevelElev is not None:
                    if bottom_z >= BotLevelElev:
                        elem.LookupParameter("Reference Level").Set(BotLevelId)

                elif BotLevelElev is None and TopLevelElev is not None:
                    if bottom_z <= TopLevelElev:
                        elem.LookupParameter("Reference Level").Set(TopLevelId)

            except Exception as e:
                TaskDialog.Show("Error", "Error processing element {}: {}".format(elem.Id, str(e)))
        t.Commit()
except Exception as e:
    TaskDialog.Show("Script Failed", "Script failed: {}".format(str(e)))
