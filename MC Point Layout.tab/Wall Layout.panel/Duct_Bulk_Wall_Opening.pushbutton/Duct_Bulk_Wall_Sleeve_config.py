import os
import clr
import System

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import (
    Window, Thickness, WindowStyle, ResizeMode,
    WindowStartupLocation, HorizontalAlignment, VerticalAlignment,
    GridLength, GridUnitType
)

from System.Windows.Controls import (
    TextBox, Button, Grid, RowDefinition, ColumnDefinition,
    ScrollViewer, StackPanel, ScrollBarVisibility, TextBlock, Orientation
)

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document

folder_name = r"c:\Temp"
filepath = os.path.join(folder_name, "Ribbon_Duct-Wall-Sleeve-Services.txt")
DEFAULT_ANNULAR = 0.5
FIRE_SMOKE_DAMPER_KEY = "Fire Smoke Damper"

if not os.path.exists(folder_name):
    os.makedirs(folder_name)


def load_service_settings(path):
    settings = {}
    if not os.path.exists(path):
        return settings

    with open(path, 'r') as f:
        for line in f:
            line = line.strip()
            if "=" not in line:
                continue

            parts = line.split("=", 1)
            if len(parts) != 2:
                continue

            key = parts[0].strip()
            val = parts[1].strip()

            try:
                settings[key] = float(val)
            except:
                pass

    return settings


def save_service_settings(path, settings):
    with open(path, 'w') as f:
        for key in sorted(settings.keys()):
            f.write("{0}={1}\n".format(key, settings[key]))


def safe_param_string(elem, param_name):
    try:
        p = elem.LookupParameter(param_name)
        if p:
            val = p.AsString()
            if val and val.strip():
                return val.strip()

            val = p.AsValueString()
            if val and val.strip():
                return val.strip()
    except:
        pass

    return None


def get_service_name(elem):
    for p_name in [
        "Fabrication Service Name",
        "Service Name",
        "System Name",
        "System Abbreviation",
        "Abbreviation"
    ]:
        val = safe_param_string(elem, p_name)
        if val:
            return val

    try:
        if hasattr(elem, "MEPSystem") and elem.MEPSystem and elem.MEPSystem.Name:
            return elem.MEPSystem.Name.strip()
    except:
        pass

    return "UNASSIGNED"


def collect_services_from_model(doc):
    services = set()

    cats = [BuiltInCategory.OST_DuctCurves]

    try:
        cats.append(BuiltInCategory.OST_FabricationDuctwork)
    except:
        pass

    for bic in cats:
        try:
            elems = (
                FilteredElementCollector(doc)
                .OfCategory(bic)
                .WhereElementIsNotElementType()
                .ToElements()
            )
            for elem in elems:
                services.add(get_service_name(elem))
        except:
            pass

    services.add(FIRE_SMOKE_DAMPER_KEY)

    return sorted(list(services), key=lambda x: (0 if x == FIRE_SMOKE_DAMPER_KEY else 1, x))


class SleeveServiceForm(Window):
    def __init__(self, services, existing_settings, default_value):
        self.Title = "Sleeve Configuration by Service"
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.result = None
        self.inputs = {}

        service_count = len(services)

        row_height = 30
        header_height = 55
        button_height = 55
        padding_height = 35

        visible_rows = min(max(service_count, 3), 12)
        list_height = visible_rows * row_height

        self.Width = 520
        self.Height = header_height + list_height + button_height + padding_height
        self.MinWidth = self.Width
        self.MinHeight = self.Height
        self.MaxHeight = header_height + (12 * row_height) + button_height + padding_height

        root = Grid()
        root.Margin = Thickness(12)
        self.Content = root

        rd0 = RowDefinition()
        rd0.Height = GridLength.Auto
        root.RowDefinitions.Add(rd0)

        rd1 = RowDefinition()
        rd1.Height = GridLength(1, GridUnitType.Star)
        root.RowDefinitions.Add(rd1)

        rd2 = RowDefinition()
        rd2.Height = GridLength.Auto
        root.RowDefinitions.Add(rd2)

        header = TextBlock()
        header.Text = "Enter annular space (inches per side) for each service:"
        header.Margin = Thickness(0, 0, 0, 10)
        Grid.SetRow(header, 0)
        root.Children.Add(header)

        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
        Grid.SetRow(scroll, 1)
        root.Children.Add(scroll)

        rows_panel = StackPanel()
        scroll.Content = rows_panel

        for service in services:
            row = Grid()
            row.Margin = Thickness(0, 0, 0, 6)

            cd0 = ColumnDefinition()
            cd0.Width = GridLength(230)
            row.ColumnDefinitions.Add(cd0)

            cd1 = ColumnDefinition()
            cd1.Width = GridLength(1, GridUnitType.Star)
            row.ColumnDefinitions.Add(cd1)

            lbl = TextBlock()
            lbl.Text = service
            lbl.VerticalAlignment = VerticalAlignment.Center
            lbl.Margin = Thickness(0, 0, 10, 0)
            Grid.SetColumn(lbl, 0)
            row.Children.Add(lbl)

            tb = TextBox()
            tb.Text = str(existing_settings.get(service, default_value))
            tb.Height = 24
            tb.Padding = Thickness(4, 2, 4, 2)
            Grid.SetColumn(tb, 1)
            row.Children.Add(tb)

            self.inputs[service] = tb
            rows_panel.Children.Add(row)

        button_panel = StackPanel()
        button_panel.Orientation = Orientation.Horizontal
        button_panel.HorizontalAlignment = HorizontalAlignment.Center
        button_panel.Margin = Thickness(0, 12, 0, 0)
        Grid.SetRow(button_panel, 2)
        root.Children.Add(button_panel)

        ok_btn = Button()
        ok_btn.Content = "OK"
        ok_btn.Width = 90
        ok_btn.Height = 28
        ok_btn.Margin = Thickness(6, 0, 6, 0)
        ok_btn.Click += self.ok_clicked
        button_panel.Children.Add(ok_btn)

        cancel_btn = Button()
        cancel_btn.Content = "Cancel"
        cancel_btn.Width = 90
        cancel_btn.Height = 28
        cancel_btn.Margin = Thickness(6, 0, 6, 0)
        cancel_btn.Click += self.cancel_clicked
        button_panel.Children.Add(cancel_btn)

        if services:
            first_service = services[0]
            self.inputs[first_service].Focus()
            self.inputs[first_service].SelectAll()

    def ok_clicked(self, sender, args):
        values = {}

        for service, tb in self.inputs.items():
            try:
                values[service] = float(tb.Text)
            except:
                tb.Focus()
                tb.SelectAll()
                return

        self.result = values
        self.DialogResult = True
        self.Close()

    def cancel_clicked(self, sender, args):
        self.DialogResult = False
        self.Close()


services_in_model = collect_services_from_model(doc)
saved_settings = load_service_settings(filepath)

form = SleeveServiceForm(services_in_model, saved_settings, DEFAULT_ANNULAR)

service_annular_map = saved_settings
if form.ShowDialog():
    if form.result:
        service_annular_map = form.result
        save_service_settings(filepath, service_annular_map)

# print(service_annular_map)