# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")
clr.AddReference("System.Xaml")

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, BuiltInParameter,
    ElementParameterFilter, FilterStringRule, ParameterValueProvider,
    FilterStringEndsWith, ElementId
)
from Autodesk.Revit.UI import TaskDialog
from System.Collections.Generic import List
from System.Windows import Window, Thickness, ResizeMode
from System.Windows.Controls import ListBox, TextBlock, Grid, RowDefinition
from System.Windows.Input import MouseButtonEventHandler
from System.Windows import SizeToContent, WindowStartupLocation

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
curview = doc.ActiveView
app = doc.Application
RevitVersion = float(app.VersionNumber)

# Utility: get parameter as value string
def get_parameter_value_by_name_AsValueString(elem, param_name):
    param = elem.LookupParameter(param_name)
    if param and param.HasValue:
        return param.AsValueString() or param.AsString()
    return ""

# Filter: Material ends with "Carbon Steel"
param_id = ElementId(BuiltInParameter.FABRICATION_PART_MATERIAL)
provider = ParameterValueProvider(param_id)
evaluator = FilterStringEndsWith()
rule = FilterStringRule(provider, evaluator, "Carbon Steel", True) if RevitVersion <= 2021 else FilterStringRule(provider, evaluator, "Carbon Steel")
material_filter = ElementParameterFilter(rule)

# Collect matching pipes
pipe_collector = FilteredElementCollector(doc, curview.Id) \
    .OfCategory(BuiltInCategory.OST_FabricationPipework) \
    .WhereElementIsNotElementType() \
    .WherePasses(material_filter) \
    .ToElements()

pipe_data = []
for pipe in pipe_collector:
    try:
        CID = pipe.ItemCustomId
        if CID == 2041:
            pipelen = pipe.get_Parameter(BuiltInParameter.FABRICATION_PART_LENGTH).AsDouble()
            pipesize = get_parameter_value_by_name_AsValueString(pipe, 'Product Entry')
            if pipelen > 20.0:
                family_name = get_parameter_value_by_name_AsValueString(pipe, 'Family')
                text = "{}: {:.2f}   DIA. {}".format(family_name, pipelen, pipesize)
                pipe_data.append((family_name, pipelen, text, pipe.Id))
    except Exception as e:
        print("Error with element: {}".format(e))

pipe_data.sort(key=lambda x: (x[0], x[1]))
pipe_data = [(item[2], item[3]) for item in pipe_data]

# WPF dialog
class PipeDialog(Window):
    def __init__(self, pipe_data):
        self.Title = "Filtered Carbon Steel Pipe List"
        self.SizeToContent = SizeToContent.WidthAndHeight
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Topmost = True
        self.ResizeMode = ResizeMode.CanResize
        self.MinWidth = 400
        self.MinHeight = 350
        self.pipe_data = pipe_data

        grid = Grid()
        grid.Margin = Thickness(10)
        grid.RowDefinitions.Add(RowDefinition())  # TextBlock row
        grid.RowDefinitions.Add(RowDefinition())  # ListBox row

        self.label = TextBlock()
        self.label.Text = "Double Click a pipe to zoom to it:"
        self.label.Margin = Thickness(0, 0, 0, 5)
        Grid.SetRow(self.label, 0)
        grid.Children.Add(self.label)

        self.listbox = ListBox()
        self.listbox.Height = 280
        for (text, eid) in pipe_data:
            self.listbox.Items.Add(text)
        self.listbox.MouseDoubleClick += MouseButtonEventHandler(self.on_double_click)
        Grid.SetRow(self.listbox, 1)
        grid.Children.Add(self.listbox)

        self.Content = grid

        # Map from text to element ID
        self.element_map = {text: eid for (text, eid) in pipe_data}

    def on_double_click(self, sender, args):
        selected = self.listbox.SelectedItem
        if selected:
            eid = self.element_map[selected]
            element = doc.GetElement(eid)
            if element:
                uidoc.Selection.SetElementIds(List[ElementId]([eid]))
                uidoc.ShowElements(eid)
                self.Close()
            else:
                TaskDialog.Show("Error", "Element not found.")

# Show the dialog
if pipe_data:
    dialog = PipeDialog(pipe_data)
    dialog.ShowDialog()
else:
    TaskDialog.Show("No Matches", "No long carbon steel pipes found.")
