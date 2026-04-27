# -*- coding: UTF-8 -*-
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import (
    Window, WindowStartupLocation, WindowStyle, GridLength,
    HorizontalAlignment, VerticalAlignment, Thickness, FontWeights
)
from System.Windows import GridUnitType
from System.Windows.Controls import (
    Grid, RowDefinition, ColumnDefinition, Button, TextBox,
    ScrollViewer, StackPanel, Orientation, Label
)
from System.Windows.Media import FontFamily

from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, FamilyInstance
)
from Autodesk.Revit.UI import TaskDialog

import string

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)


# ── helpers ──────────────────────────────────────────────────────────────────

def get_id_value(id_obj):
    if id_obj is None:
        return None
    try:
        return id_obj.Value if RevitINT > 2025 else id_obj.IntegerValue
    except:
        try:
            return id_obj.Value
        except:
            return id_obj.IntegerValue


def get_element_id_value(element):
    if not element:
        return None
    try:
        return get_id_value(element.Id)
    except:
        return None


def is_fab_hanger(element):
    try:
        return element and element.Category and \
               get_id_value(element.Category.Id) == int(BuiltInCategory.OST_FabricationHangers)
    except:
        return False


def is_trapeze_hanger(element):
    if not is_fab_hanger(element):
        return False
    try:
        rod_info = element.GetRodInfo()
        return rod_info is not None and rod_info.RodCount > 1
    except:
        return False


def is_beam_hanger(element):
    if not is_fab_hanger(element):
        return False
    try:
        p = element.LookupParameter("FP_Beam Hanger")
        if p and p.HasValue:
            val = p.AsString()
            if val and val.strip().lower() == "yes":
                return True
    except:
        pass
    return False


def is_gtp_element(element):
    if not element or not element.Category:
        return False
    if get_id_value(element.Category.Id) != int(BuiltInCategory.OST_GenericModel):
        return False
    try:
        symbol = element.Symbol
        if symbol:
            if getattr(symbol, 'FamilyName', '') == "GTP":
                return True
            if hasattr(symbol, 'Family') and symbol.Family and symbol.Family.Name == "GTP":
                return True
    except:
        pass
    return False


def is_seismic_brace_gtp(element):
    if not is_gtp_element(element):
        return False
    try:
        if isinstance(element, FamilyInstance) and element.SuperComponent:
            parent = element.SuperComponent
            if parent and parent.Category:
                if get_id_value(parent.Category.Id) == int(BuiltInCategory.OST_StructuralStiffener):
                    return True
    except:
        pass
    return False


def get_suffix(index):
    if index < 26:
        return string.ascii_uppercase[index]
    first = (index // 26) - 1
    second = index % 26
    return string.ascii_uppercase[first] + string.ascii_uppercase[second]


def make_child_point_number(base_id, index):
    return "{}_{}".format(base_id, get_suffix(index))


def make_rck_child_point_number(base_id, index):
    return "{}{}".format(base_id, get_suffix(index))


def get_parameter_value(element, param_name):
    if not element:
        return None
    try:
        p = element.LookupParameter(param_name)
        if p and p.HasValue:
            val = p.AsString()
            if val and val.strip():
                return val.strip()
    except:
        pass
    return None


def clean_z(val, tol=1e-9):
    try:
        return 0.0 if abs(val) < tol else val
    except:
        return val


def get_shared_coords(point):
    """Convert internal point to shared N, E, Elev."""
    if point is None:
        return None, None, None
    try:
        proj_pos = doc.ActiveProjectLocation.GetProjectPosition(point)
        north = proj_pos.NorthSouth
        east = proj_pos.EastWest
        elev = clean_z(proj_pos.Elevation)
        return north, east, elev
    except:
        return point.X, point.Y, clean_z(point.Z)


def get_ts_point_number(element):
    """Read TS_Point_Number directly — no renumbering."""
    val = get_parameter_value(element, "TS_Point_Number")
    return val if val else ""


# ── collect elements ──────────────────────────────────────────────────────────

def collect_elements():
    hangers = list(
        FilteredElementCollector(doc, curview.Id)
        .OfCategory(BuiltInCategory.OST_FabricationHangers)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    generics = list(
        FilteredElementCollector(doc, curview.Id)
        .OfCategory(BuiltInCategory.OST_GenericModel)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    gtps = [el for el in generics if is_gtp_element(el)]

    all_elements = list(hangers) + gtps

    # Apply exclusions
    filtered = [
        el for el in all_elements
        if not (is_fab_hanger(el) and is_beam_hanger(el))
        and not is_seismic_brace_gtp(el)
    ]

    return sorted(filtered, key=lambda el: get_element_id_value(el))


# ── build point lines ─────────────────────────────────────────────────────────

def build_point_lines(elements):
    lines = []

    for element in elements:
        point_number = get_ts_point_number(element)
        if not point_number:
            continue  # skip elements with no point number assigned

        if is_fab_hanger(element):
            try:
                rod_info = element.GetRodInfo()
                if rod_info is None:
                    continue
                rod_count = rod_info.RodCount

                for i in range(rod_count):
                    pos = rod_info.GetRodEndPosition(i)
                    if pos is None:
                        continue

                    if rod_count == 1:
                        pid = point_number
                    elif is_trapeze_hanger(element):
                        pid = make_rck_child_point_number(point_number, i)
                    else:
                        pid = make_child_point_number(point_number, i)

                    north, east, elev = get_shared_coords(pos)
                    if north is None:
                        continue

                    lines.append("POINT {},{},{} ".format(
                        round(east * 12, 8),
                        round(north * 12, 8),
                        round(elev * 12, 8)
                    ))
            except:
                continue

        else:
            # GTP
            # Check parent for point number override
            owner_number = point_number
            if isinstance(element, FamilyInstance) and element.SuperComponent:
                parent_num = get_ts_point_number(element.SuperComponent)
                if parent_num:
                    owner_number = parent_num

            try:
                loc = element.Location
                if hasattr(loc, 'Point') and loc.Point is not None:
                    pos = loc.Point
                else:
                    pos = element.GetTransform().Origin

                north, east, elev = get_shared_coords(pos)
                if north is None:
                    continue

                lines.append("POINT {},{},{} ".format(
                    round(east * 12, 8),
                    round(north * 12, 8),
                    round(elev * 12, 8)
                ))
            except:
                continue

    return ["SETAPLCOLOR "] + lines


# ── display window ────────────────────────────────────────────────────────────

class PointDisplayWindow(Window):
    def __init__(self, lines):
        self.Title = "Point Export — Copy to CAD"
        self.Width = 680
        self.Height = 520
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.ResizeMode = 0  # CanResizeWithGrip
        self.Topmost = True

        root = Grid()
        root.Margin = Thickness(10)
        self.Content = root

        root.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))   # header
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))  # text
        root.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))   # buttons

        # Header label
        header = Label()
        header.Content = "{} point(s) — Copy, then paste into CAD command line".format(len(lines))
        header.FontWeight = FontWeights.Bold
        header.Margin = Thickness(0, 0, 0, 6)
        Grid.SetRow(header, 0)
        root.Children.Add(header)

        # Scrollable text box
        scroll = ScrollViewer()
        Grid.SetRow(scroll, 1)
        root.Children.Add(scroll)

        self.txt = TextBox()
        self.txt.FontFamily = FontFamily("Consolas")
        self.txt.FontSize = 12
        self.txt.IsReadOnly = True
        self.txt.AcceptsReturn = True
        self.txt.VerticalScrollBarVisibility = 0   # Auto
        self.txt.HorizontalScrollBarVisibility = 0
        self.txt.Text = "\r\n".join(lines)
        scroll.Content = self.txt

        # Buttons
        btn_panel = StackPanel()
        btn_panel.Orientation = Orientation.Horizontal
        btn_panel.HorizontalAlignment = HorizontalAlignment.Right
        btn_panel.Margin = Thickness(0, 8, 0, 0)
        Grid.SetRow(btn_panel, 2)
        root.Children.Add(btn_panel)

        btn_copy = Button()
        btn_copy.Content = "Copy"
        btn_copy.Width = 90
        btn_copy.Height = 28
        btn_copy.Margin = Thickness(0, 0, 8, 0)
        btn_copy.Click += self.on_copy
        btn_panel.Children.Add(btn_copy)

        btn_close = Button()
        btn_close.Content = "Close"
        btn_close.Width = 90
        btn_close.Height = 28
        btn_close.Click += lambda s, e: self.Close()
        btn_panel.Children.Add(btn_close)

    def on_copy(self, sender, args):
        self.txt.SelectAll()
        self.txt.Copy()


# ── main ──────────────────────────────────────────────────────────────────────

elements = collect_elements()

if not elements:
    TaskDialog.Show("No Elements", "No fabrication hangers or GTP points found in the active view.")
else:
    lines = build_point_lines(elements)

    if not lines:
        TaskDialog.Show("No Points", "Elements were found but none had a TS_Point_Number assigned.")
    else:
        win = PointDisplayWindow(lines)
        win.ShowDialog()