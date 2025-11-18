from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, BuiltInParameter,
    ElementParameterFilter, FilterStringRule, ParameterValueProvider, FilterStringEndsWith,
    ElementId
)
from Autodesk.Revit.UI import TaskDialog
from System.Collections.Generic import List

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
from System.Windows import Application, Window, Thickness, WindowStyle, ResizeMode, WindowStartupLocation, GridLength
from System.Windows.Controls import Label, ListBox, Grid, RowDefinition
from System.Windows.Media import Brushes

import System
from System import Action
import System.Windows.Threading

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
curview = doc.ActiveView
app = doc.Application
RevitVersion = float(app.VersionNumber)

# --- External param utility function ---
def get_parameter_value_by_name_AsValueString(elem, param_name):
    param = elem.LookupParameter(param_name)
    if param and param.HasValue:
        return param.AsValueString() or param.AsString()
    return ""

# --- Filter for pipes with material "Cast Iron" ---
param_id = ElementId(BuiltInParameter.FABRICATION_PART_MATERIAL)
provider = ParameterValueProvider(param_id)
evaluator = FilterStringEndsWith()
if RevitVersion <= 2021:
    rule = FilterStringRule(provider, evaluator, "Cast Iron", True)
else:
    rule = FilterStringRule(provider, evaluator, "Cast Iron")
material_filter = ElementParameterFilter(rule)

pipe_collector = FilteredElementCollector(doc, curview.Id) \
    .OfCategory(BuiltInCategory.OST_FabricationPipework) \
    .WhereElementIsNotElementType() \
    .WherePasses(material_filter) \
    .ToElements()
epsilon = 1e-3
# --- Collect matching elements ---
pipe_data = []
for pipe in pipe_collector:
    try:
        CID = pipe.ItemCustomId
        if CID == 2041:
            pipelen = pipe.get_Parameter(BuiltInParameter.FABRICATION_PART_LENGTH).AsDouble()
            pipesize = get_parameter_value_by_name_AsValueString(pipe, 'Product Entry')
            if (pipelen > 10.0 or (pipelen >= 9.9789 and pipelen < 10.0)) and abs(pipelen - 10.0) >= epsilon:
                family_name = get_parameter_value_by_name_AsValueString(pipe, 'Family')
                text = "{}: {:.3f}   DIA. {}".format(family_name, pipelen, pipesize)
                pipe_data.append((family_name, pipelen, text, pipe.Id))
    except Exception as e:
        print("Error with element: {}".format(e))

# Sort pipe_data by family_name (index 0), then pipelen (index 1)
pipe_data.sort(key=lambda x: (x[0], x[1]))

# Update pipe_data to keep only (text, pipe.Id) after sorting
pipe_data = [(item[2], item[3]) for item in pipe_data]

# --- WPF Window ---
class PipeListForm(Window):
    def __init__(self, pipe_data, doc, uidoc):
        self.Title = "Filtered Pipe List"
        self.Width = 400
        self.Height = 400
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.ResizeMode = ResizeMode.CanResize
        self.Topmost = True

        self.doc = doc
        self.uidoc = uidoc

        grid = Grid()
        grid.Margin = Thickness(10)

        # Define layout: label + tight spacer + listbox
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))  # Label row
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(2)))    # Spacer
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))  # ListBox row

        self.Content = grid

        label = Label()
        label.Content = "Double Click a pipe to zoom to it:"
        label.Margin = Thickness(0)
        label.Foreground = Brushes.Black
        Grid.SetRow(label, 0)
        grid.Children.Add(label)

        self.listbox = ListBox()
        self.listbox.Margin = Thickness(0)
        self.listbox.Height = 300
        for (text, eid) in pipe_data:
            self.listbox.Items.Add(text)
        Grid.SetRow(self.listbox, 2)
        grid.Children.Add(self.listbox)

        # Map text to element id
        self.element_map = {text: eid for (text, eid) in pipe_data}
        self.listbox.MouseDoubleClick += self.select_element

    def select_element(self, sender, args):
        selected_text = self.listbox.SelectedItem
        if selected_text:
            eid = self.element_map[selected_text]
            element = self.doc.GetElement(eid)
            if element:
                self.uidoc.Selection.SetElementIds(List[ElementId]([eid]))
                self.uidoc.ShowElements(eid)
                self.Close()
            else:
                TaskDialog.Show("Error", "Element not found.")

# --- Launch UI ---
if pipe_data:
    form = PipeListForm(pipe_data, doc, uidoc)
    form.Show()
    while form.IsVisible:
        Dispatcher = System.Windows.Threading.Dispatcher.CurrentDispatcher
        Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Background, Action(lambda: None))
else:
    TaskDialog.Show("No Matches", "No long pipes found.")
