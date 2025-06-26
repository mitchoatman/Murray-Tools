# -*- coding: utf-8 -*-
import clr

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, BuiltInParameter,
    ElementParameterFilter, FilterStringRule, ParameterValueProvider, FilterStringEndsWith,
    ElementId
)
from Autodesk.Revit.UI import TaskDialog

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Windows.Forms import Application, Form, ListBox, Label, FormStartPosition, FormBorderStyle
from System.Drawing import Point, Size
from System.Collections.Generic import List

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

# --- Filter for pipes with material "Cast Iron:" ---
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

# --- Collect matching elements ---
pipe_data = []
for pipe in pipe_collector:
    try:
        CID = pipe.ItemCustomId
        if CID == 2041:
            pipelen = pipe.get_Parameter(BuiltInParameter.FABRICATION_PART_LENGTH).AsDouble()
            pipesize = get_parameter_value_by_name_AsValueString(pipe, 'Product Entry')
            if pipelen > 10.002:
                family_name = get_parameter_value_by_name_AsValueString(pipe, 'Family')
                text = "{}: {:.2f}   DIA. {}".format(family_name, pipelen, pipesize)
                pipe_data.append((family_name, pipelen, text, pipe.Id))
    except Exception as e:
        print("Error with element: {}".format(e))

# Sort pipe_data by family_name (index 0), then pipelen (index 1)
pipe_data.sort(key=lambda x: (x[0], x[1]))

# Update pipe_data to keep only (text, pipe.Id) after sorting
pipe_data = [(item[2], item[3]) for item in pipe_data]

# --- WinForm UI ---
class PipeListForm(Form):
    def __init__(self, pipe_data, doc, uidoc):
        self.Text = "Filtered Pipe List"
        self.Size = Size(400, 400)
        self.StartPosition = FormStartPosition.CenterScreen
        self.TopMost = True
        self.ShowIcon = False
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.doc = doc  # Store doc
        self.uidoc = uidoc  # Store uidoc

        self.label = Label()
        self.label.Text = "Double Click a pipe to zoom to it:"
        self.label.Location = Point(10, 10)
        self.label.Size = Size(380, 20)
        self.Controls.Add(self.label)

        self.listbox = ListBox()
        self.listbox.Location = Point(10, 40)
        self.listbox.Size = Size(365, 300)
        for (text, eid) in pipe_data:
            self.listbox.Items.Add(text)
        self.Controls.Add(self.listbox)

        # Store mapping of text to element ID
        self.element_map = {text: eid for (text, eid) in pipe_data}
        self.listbox.DoubleClick += self.select_element

    def select_element(self, sender, args):
        selected_text = self.listbox.SelectedItem
        if selected_text:
            eid = self.element_map[selected_text]
            element = self.doc.GetElement(eid)
            if element:
                self.uidoc.Selection.SetElementIds(List[ElementId]([eid]))
                self.uidoc.ShowElements(eid)
                self.Close()  # Close the dialog after selecting and zooming
            else:
                TaskDialog.Show("Error", "Element not found.")

# --- Show form or fallback message ---
if pipe_data:
    form = PipeListForm(pipe_data, doc, uidoc)
    form.Show()  # Non-modal display
    while form.Visible:
        Application.DoEvents()
else:
    TaskDialog.Show("No Matches", "No long pipes found.")