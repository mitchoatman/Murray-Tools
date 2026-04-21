from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, BuiltInParameter,
    ElementParameterFilter, FilterStringRule, ParameterValueProvider, FilterStringEndsWith,
    ElementId, Transaction, FabricationPart
)
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference("System.Core")
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from System.Collections.Generic import List, HashSet
from System.Windows import Application, Window, Thickness, WindowStyle, ResizeMode, WindowStartupLocation, GridLength
from System.Windows.Controls import Label, ListBox, Button, Grid, RowDefinition
from System.Windows.Media import Brushes

import System
from System import Action
import System.Windows.Threading

# Get the active document and UI document
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
curview = doc.ActiveView
app = doc.Application
RevitVersion = float(app.VersionNumber)

# Utility to read string parameters
def get_param_str(elem, name):
    p = elem.LookupParameter(name)
    if p and p.HasValue:
        return p.AsValueString() or p.AsString()
    return ""

# Build the same cast-iron/length filter and collect pipe_data
def collect_pipes():
    param_id = ElementId(BuiltInParameter.FABRICATION_PART_MATERIAL)
    prov = ParameterValueProvider(param_id)
    evalr = FilterStringEndsWith()
    rule = FilterStringRule(prov, evalr, "Cast Iron", True) if RevitVersion <= 2021 \
           else FilterStringRule(prov, evalr, "Cast Iron")
    material_filt = ElementParameterFilter(rule)

    pipes = FilteredElementCollector(doc, curview.Id) \
        .OfCategory(BuiltInCategory.OST_FabricationPipework) \
        .WhereElementIsNotElementType() \
        .WherePasses(material_filt) \
        .ToElements()

    out = []
    eps = 1e-3
    for p in pipes:
        try:
            if p.ItemCustomId == 2041:
                L = p.get_Parameter(BuiltInParameter.FABRICATION_PART_LENGTH).AsDouble()
                D = get_param_str(p, 'Product Entry')
                if (L > 10.0 or (L >= 9.9789 and L < 10.0)) and abs(L - 10.0) >= eps:
                    fam = get_param_str(p, 'Family')
                    txt = "{}: {:.3f}   DIA. {}".format(fam, L, D)
                    out.append((fam, L, txt, p.Id))
        except:
            pass
    out.sort(key=lambda x: (x[0], x[1]))
    return [(t, eid) for (_, _, t, eid) in out]

# The WPF form
class PipeListForm(Window):
    def __init__(self, pipe_data):
        Window.__init__(self)
        self.Title = "Filtered Pipe List"
        self.Width = 400
        self.Height = 415
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.ResizeMode = ResizeMode.CanResize
        self.Topmost = True

        self.pipe_data = pipe_data  # store for refresh
        self.doc = doc
        self.uidoc = uidoc

        g = Grid(); g.Margin = Thickness(10)
        g.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        g.RowDefinitions.Add(RowDefinition(Height=GridLength(2)))
        g.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        g.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        self.Content = g

        lbl = Label(Content="Double-click to zoom, or click Update/Optimize:")
        lbl.Foreground = Brushes.Black
        Grid.SetRow(lbl, 0); g.Children.Add(lbl)

        self.lb = ListBox(Height=300)
        Grid.SetRow(self.lb, 2); g.Children.Add(self.lb)
        self.lb.MouseDoubleClick += self.on_double

        btn = Button(Content="Update/Optimize")
        btn.Click += self.on_click
        Grid.SetRow(btn, 3); g.Children.Add(btn)

        self.populate(pipe_data)

    def populate(self, data):
        self.lb.Items.Clear()
        self.map = {}
        for txt, eid in data:
            self.lb.Items.Add(txt)
            self.map[txt] = eid

    def on_double(self, s, e):
        txt = self.lb.SelectedItem
        if not txt: return
        eid = self.map[txt]
        self.uidoc.Selection.SetElementIds(List[ElementId]([eid]))
        self.uidoc.ShowElements(eid)

    def on_click(self, s, e):
        # optimize selected
        txt = self.lb.SelectedItem
        if txt:
            eid = self.map[txt]
            self.optimize_one(eid)
        # then refresh list
        self.pipe_data = collect_pipes()
        self.populate(self.pipe_data)

    def optimize_one(self, eid):
        try:
            ids = HashSet[ElementId](); ids.Add(eid)
            t = Transaction(self.doc, "Optimize lengths")
            t.Start()
            optimized = FabricationPart.OptimizeLengths(self.doc, ids)
            t.Commit()
            # TaskDialog.Show("Optimization", 
                # "Optimization complete for {0} straight parts.".format(len(optimized)))
        except Exception as ex:
            TaskDialog.Show("Error", "Optimization failed: {0}".format(str(ex)))

# --- Launch ---
data = collect_pipes()
if data:
    w = PipeListForm(data)
    w.Show()
    # keep responsive
    disp = System.Windows.Threading.Dispatcher.CurrentDispatcher
    while w.IsVisible:
        disp.Invoke(System.Windows.Threading.DispatcherPriority.Background, Action(lambda: None))
else:
    TaskDialog.Show("No Matches", "No long pipes found.")