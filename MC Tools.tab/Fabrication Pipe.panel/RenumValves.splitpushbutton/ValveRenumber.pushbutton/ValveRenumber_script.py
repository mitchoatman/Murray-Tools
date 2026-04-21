import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter, FilterStringRule, \
    ParameterValueProvider, ElementId, Transaction, FilterStringEquals, ElementParameterFilter, FabricationPart
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.UI import TaskDialog
from Parameters.Add_SharedParameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString
import os, clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
import System
from System.Windows import Application, Window, Thickness, WindowStyle, ResizeMode, WindowStartupLocation, HorizontalAlignment, VerticalAlignment
from System.Windows.Controls import Label, TextBox, Button, Grid, RowDefinition, CheckBox, ScrollViewer, StackPanel, ColumnDefinition
from System.Collections.Generic import List

Shared_Params()
DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

folder_name = "c:\\Temp"
filepath = os.path.join(folder_name, 'Ribbon_ValveRENumber.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)

if not os.path.exists(filepath):
    with open(filepath, 'w') as f:
        f.write('1')

with open(filepath, 'r') as f:
    PrevInput = f.read().strip()

# Filter creation helpers
def create_filter_2023_newer(key_parameter, element_value):
    f_parameter = ParameterValueProvider(ElementId(key_parameter))
    f_rule = FilterStringRule(f_parameter, FilterStringEquals(), element_value)
    return ElementParameterFilter(f_rule)

def create_filter_2022_older(key_parameter, element_value):
    f_parameter = ParameterValueProvider(ElementId(key_parameter))
    f_rule = FilterStringRule(f_parameter, FilterStringEquals(), element_value, False)
    return ElementParameterFilter(f_rule)

def set_parameter_by_name(element, parameterName, value):
    p = element.LookupParameter(parameterName)
    if p and p.StorageType == DB.StorageType.String:
        p.Set(value)

def set_customdata_by_custid(fabpart, custid, value):
    fabpart.SetPartCustomDataText(custid, value)

# Original starting number dialog (unchanged)
class ValveNumberForm(Window):
    def __init__(self, default_value):
        self.Title = "Valve Number"
        self.Width = 300
        self.Height = 150
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.result = None
        # Grid layout
        grid = Grid()
        grid.Margin = Thickness(10)
        row1 = RowDefinition()
        row2 = RowDefinition()
        row3 = RowDefinition()
        grid.RowDefinitions.Add(row1)
        grid.RowDefinitions.Add(row2)
        grid.RowDefinitions.Add(row3)
        self.Content = grid
        # Label
        label = Label()
        label.Content = "Enter Starting Valve Number:"
        label.Margin = Thickness(0, -5, 0, 10)
        Grid.SetRow(label, 0)
        grid.Children.Add(label)
        # TextBox
        self.textbox = TextBox()
        self.textbox.Text = default_value
        self.textbox.Margin = Thickness(0, 0, 0, 10)
        Grid.SetRow(self.textbox, 1)
        grid.Children.Add(self.textbox)
        # OK Button (centered)
        ok_button = Button()
        ok_button.Content = "OK"
        ok_button.Width = 75
        ok_button.HorizontalAlignment = HorizontalAlignment.Center
        ok_button.Click += self.on_ok
        Grid.SetRow(ok_button, 2)
        grid.Children.Add(ok_button)

    def OnContentRendered(self, e):
        Window.OnContentRendered(self, e)
        self.textbox.Focus()
        self.textbox.SelectAll()

    def on_ok(self, sender, args):
        self.result = self.textbox.Text
        self.DialogResult = True
        self.Close()

# Exclusion dialog (X button remains; no reliable way to remove it without losing title bar)
class AliasExclusionDialog(Window):
    def __init__(self, all_aliases):
        self.Title = "Select Aliases to EXCLUDE"
        self.Width = 380
        self.Height = 380
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.result = None

        grid = Grid()
        grid.Margin = Thickness(10)
        grid.RowDefinitions.Add(RowDefinition(Height=System.Windows.GridLength.Auto))
        grid.RowDefinitions.Add(RowDefinition(Height=System.Windows.GridLength(1, System.Windows.GridUnitType.Star)))
        grid.RowDefinitions.Add(RowDefinition(Height=System.Windows.GridLength.Auto))

        # Header label
        lbl = Label()
        lbl.Content = "Check the aliases you want to skip (not renumber):"
        lbl.Margin = Thickness(0, 0, 0, 12)
        Grid.SetRow(lbl, 0)
        grid.Children.Add(lbl)

        # Scrollable checkbox list
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto
        stack = StackPanel()
        stack.Margin = Thickness(4)

        self.checkboxes = {}
        common_excluded = {'STRAINER', 'CHECK', 'BALANCE', 'PLUG'}
        display_aliases = sorted(set(all_aliases) | common_excluded)

        for alias in display_aliases:
            if not alias:
                continue
            cb = CheckBox()
            cb.Content = alias
            cb.IsChecked = alias in common_excluded
            cb.Margin = Thickness(0, 4, 0, 4)
            stack.Children.Add(cb)
            self.checkboxes[alias] = cb

        scroll.Content = stack
        Grid.SetRow(scroll, 1)
        grid.Children.Add(scroll)

        # OK button - identical to ValveNumberForm
        ok_button = Button()
        ok_button.Content = "OK"
        ok_button.Width = 75
        ok_button.HorizontalAlignment = HorizontalAlignment.Center
        ok_button.Click += self.on_ok
        Grid.SetRow(ok_button, 2)
        grid.Children.Add(ok_button)

        self.Content = grid

    def on_ok(self, sender, args):
        excluded = set()
        for alias, cb in self.checkboxes.items():
            if cb.IsChecked == True:
                excluded.add(alias)
        self.result = excluded
        self.DialogResult = True
        self.Close()

# Selection filter (unchanged)
class ValveSelectionFilter(ISelectionFilter):
    def AllowElement(self, e):
        if not isinstance(e, FabricationPart):
            return False
        return e.ServiceType == 53

# Renumbering logic (unchanged)
def renumber_valves_by_proximity(selected_valve, excluded_aliases):
    collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationPipework)
    valves_to_renumber = [el for el in collector.ToElements()
                          if el.ServiceType == 53 and el.Alias not in excluded_aliases]
    if not valves_to_renumber:
        TaskDialog.Show("No Valves Found", "No eligible valves found in current view.")
        return False
    # Read previous input
    try:
        with open(filepath, 'r') as f:
            PrevInput = f.read().strip()
    except:
        PrevInput = "1"
    import re
    def extract_number_and_prefix(val):
        match = re.search(r'(\d+)$', val)
        if match:
            num = match.group(1)
            prefix = val[:match.start()]
            return prefix, int(num), len(num)
        else:
            return val, 0, 0
    prefix, num, num_length = extract_number_and_prefix(PrevInput)
    initial_value = "{}{}".format(prefix, str(num + 1).zfill(num_length or 1))
    form = ValveNumberForm(initial_value)
    if form.ShowDialog() and form.DialogResult:
        value = form.result
    else:
        return False
    with open(filepath, 'w') as f:
        f.write(value)
    prefix, start_num, num_length = extract_number_and_prefix(value)
    num_length = num_length or 1
    def distance_between_parts(part1, part2):
        point1 = part1.Origin
        point2 = part2.Origin
        return point1.DistanceTo(point2)
    sorted_valves = sorted(valves_to_renumber, key=lambda x: distance_between_parts(selected_valve, x))
    numincrement = start_num - 1
    last_number = None
    for valve in sorted_valves:
        numincrement += 1
        newvalvenumber = "{}{}".format(prefix, str(numincrement).zfill(num_length))
        set_parameter_by_name(valve, 'FP_Valve Number', newvalvenumber)
        set_customdata_by_custid(valve, 2, newvalvenumber)
        last_number = newvalvenumber
    if last_number:
        with open(filepath, 'w') as f:
            f.write(last_number)
    return True

# Main logic (unchanged)
t = Transaction(doc, 'Set Valve Number and Isolate Service')
t.Start()

from Autodesk.Revit.Exceptions import OperationCanceledException

try:
    selected_part_ref = uidoc.Selection.PickObject(ObjectType.Element, ValveSelectionFilter(), "Select Valve to start numbering from")
    selected_valve = doc.GetElement(selected_part_ref)
except OperationCanceledException:
    t.RollBack()
    TaskDialog.Show("Cancelled", "Operation cancelled by user.")
    import sys
    sys.exit()
except Exception as e:
    t.RollBack()
    TaskDialog.Show("Error", str(e))
    import sys
    sys.exit()

# Collect aliases in view for the exclusion dialog
all_parts = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationPipework).ToElements()
aliases_in_view = {p.Alias for p in all_parts if p.Alias and p.ServiceType == 53}

# Show exclusion dialog
exclusion_dlg = AliasExclusionDialog(aliases_in_view)
if not exclusion_dlg.ShowDialog() or exclusion_dlg.result is None:
    t.RollBack()
    TaskDialog.Show("Cancelled", "Operation cancelled by user.")
    import sys
    sys.exit()

excluded_aliases = exclusion_dlg.result

# Optional: isolate same service
service_name = get_parameter_value_by_name_AsString(selected_valve, 'Fabrication Service Name')
if service_name:
    if RevitINT > 2022:
        service_filter = create_filter_2023_newer(BuiltInParameter.FABRICATION_SERVICE_NAME, service_name)
    else:
        service_filter = create_filter_2022_older(BuiltInParameter.FABRICATION_SERVICE_NAME, service_name)
    analyticalCollector = FilteredElementCollector(doc).WherePasses(service_filter).ToElementIds()
    curview.IsolateElementsTemporary(analyticalCollector)

# Renumber
success = renumber_valves_by_proximity(selected_valve, excluded_aliases)

curview.DisableTemporaryViewMode(DB.TemporaryViewMode.TemporaryHideIsolate)

if success:
    t.Commit()
else:
    t.RollBack()