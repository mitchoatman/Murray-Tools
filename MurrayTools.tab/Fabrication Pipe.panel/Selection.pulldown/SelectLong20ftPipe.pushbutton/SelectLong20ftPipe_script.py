from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, BuiltInParameter,
    ElementId, Transaction, FabricationPart
)
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference("System.Core")
from Autodesk.Revit.UI import TaskDialog
from System.Collections.Generic import List, HashSet
from System.Windows import Application, Window, Thickness, WindowStyle, ResizeMode, WindowStartupLocation, GridLength
from System.Windows.Controls import Label, ListBox, Button, Grid, RowDefinition
from System.Windows.Media import Brushes

import System
from System import Action
import System.Windows.Threading

# Revit contexts
uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document
curview = doc.ActiveView
app   = doc.Application
RevitVersion = float(app.VersionNumber)

# Utility to get a string parameter
def get_parameter_value_by_name_AsValueString(elem, param_name):
    p = elem.LookupParameter(param_name)
    if p and p.HasValue:
        return p.AsValueString() or p.AsString()
    return ""

# Collect pipes of any of the three materials, over length 20.002
# Collect pipes of any of the three materials, over length 20.002
def collect_pipes():
    wanted_materials = ["copper", "carbon steel", "pvc"]
    all_pipes = FilteredElementCollector(doc, curview.Id) \
        .OfCategory(BuiltInCategory.OST_FabricationPipework) \
        .WhereElementIsNotElementType() \
        .ToElements()

    out = []
    for p in all_pipes:
        try:
            if p.ItemCustomId != 2041:
                continue

            # read and normalize the material value
            matParam = p.get_Parameter(BuiltInParameter.FABRICATION_PART_MATERIAL)
            raw = matParam.AsValueString() or matParam.AsString() or ""
            matVal = raw.strip().lower()

            # skip if not one of our three
            if not any(matVal.find(m) != -1 for m in wanted_materials):
                continue

            # length check
            L = p.get_Parameter(BuiltInParameter.FABRICATION_PART_LENGTH).AsDouble()
            if L <= 20.002:
                continue

            # build display text
            D   = get_parameter_value_by_name_AsValueString(p, 'Product Entry')
            fam = get_parameter_value_by_name_AsValueString(p, 'Family')
            txt = "{}: {:.2f}   DIA. {}".format(fam, L, D)

            out.append((fam, L, txt, p.Id))
        except Exception:
            # swallow any element that misbehaves
            pass

    out.sort(key=lambda x: (x[0], x[1]))
    return [(txt, eid) for (_, _, txt, eid) in out]

# WPF dialog exactly as before
class PipeListForm(Window):
    def __init__(self, data):
        Window.__init__(self)
        self.Title = "Filtered Pipe List"
        self.Width = 400
        self.Height = 415
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.ResizeMode = ResizeMode.CanResize
        self.Topmost = True

        self.doc = doc
        self.uidoc = uidoc
        self.pipe_data = data

        grid = Grid()
        grid.Margin = Thickness(10)
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength(2)))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        self.Content = grid

        lbl = Label(Content="Double-click to zoom:")
        lbl.Foreground = Brushes.Black
        Grid.SetRow(lbl, 0)
        grid.Children.Add(lbl)

        self.lb = ListBox(Height=300)
        self.lb.MouseDoubleClick += self.on_double
        Grid.SetRow(self.lb, 2)
        grid.Children.Add(self.lb)

        btn = Button(Content="Update/Optimize")
        btn.Click += self.on_click
        Grid.SetRow(btn, 3)
        grid.Children.Add(btn)

        self.populate(self.pipe_data)

    def populate(self, data):
        self.lb.Items.Clear()
        self.map = {}
        for txt, eid in data:
            self.lb.Items.Add(txt)
            self.map[txt] = eid

    def on_double(self, sender, args):
        txt = self.lb.SelectedItem
        if not txt: return
        eid = self.map[txt]
        self.uidoc.Selection.SetElementIds(List[ElementId]([eid]))
        self.uidoc.ShowElements(eid)

    def on_click(self, sender, args):
        txt = self.lb.SelectedItem
        if txt:
            eid = self.map[txt]
            self.optimize_one(eid)
        self.pipe_data = collect_pipes()
        self.populate(self.pipe_data)

    def optimize_one(self, eid):
        try:
            ids = HashSet[ElementId]()
            ids.Add(eid)
            t = Transaction(self.doc, "Optimize lengths")
            t.Start()
            FabricationPart.OptimizeLengths(self.doc, ids)
            t.Commit()
            # you can uncomment to show a TaskDialog:
            # TaskDialog.Show("Optimization", "Done.")
        except Exception as ex:
            TaskDialog.Show("Error", "Optimization failed: {0}".format(str(ex)))

# Launch
data = collect_pipes()
if data:
    w = PipeListForm(data)
    w.Show()
    disp = System.Windows.Threading.Dispatcher.CurrentDispatcher
    while w.IsVisible:
        disp.Invoke(System.Windows.Threading.DispatcherPriority.Background, Action(lambda: None))
else:
    TaskDialog.Show("No Matches", "No long pipes found.")