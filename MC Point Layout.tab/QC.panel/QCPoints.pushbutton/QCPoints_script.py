# -*- coding: UTF-8 -*-
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import (
    Window, WindowStartupLocation, WindowStyle, GridLength,
    HorizontalAlignment, Thickness, FontWeights
)
from System.Windows import GridUnitType
from System.Windows.Controls import (
    Grid, RowDefinition, Button, TextBox,
    ScrollViewer, StackPanel, Orientation, Label
)
from System.Windows.Media import FontFamily

from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory, FamilyInstance
)
from Autodesk.Revit.UI import TaskDialog

import string
from collections import defaultdict

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

    try:
        if hasattr(element, "Symbol") and element.Symbol:
            p = element.Symbol.LookupParameter(param_name)
            if p and p.HasValue:
                val = p.AsString()
                if val and val.strip():
                    return val.strip()
    except:
        pass

    try:
        type_id = element.GetTypeId()
        if type_id != DB.ElementId.InvalidElementId:
            elem_type = doc.GetElement(type_id)
            if elem_type:
                p = elem_type.LookupParameter(param_name)
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
    if point is None:
        return None, None, None
    try:
        proj_pos = doc.ActiveProjectLocation.GetProjectPosition(point)
        north = proj_pos.NorthSouth
        east = proj_pos.EastWest
        elev = clean_z(proj_pos.Elevation)
        return north, east, elev
    except:
        return point.Y, point.X, clean_z(point.Z)


def get_ts_point_number(element):
    val = get_parameter_value(element, "TS_Point_Number")
    return val if val else ""


def get_point_number_owner(element):
    try:
        if is_gtp_element(element) and isinstance(element, FamilyInstance) and element.SuperComponent:
            return element.SuperComponent
    except:
        pass
    return element


def get_all_gtps_from_element(element):
    gtps = []

    if is_gtp_element(element):
        gtps.append(element)

    if isinstance(element, FamilyInstance):
        try:
            if hasattr(element, "GetSubComponentIds"):
                for sub_id in element.GetSubComponentIds():
                    sub_el = doc.GetElement(sub_id)
                    if sub_el:
                        gtps.extend(get_all_gtps_from_element(sub_el))

            if hasattr(element, "GetSubelements"):
                for sub in element.GetSubelements():
                    sub_el = doc.GetElement(sub.ElementId)
                    if sub_el:
                        gtps.extend(get_all_gtps_from_element(sub_el))

            if hasattr(element, "GetDependentElements"):
                dep_ids = element.GetDependentElements(None)
                for dep_id in dep_ids:
                    dep_el = doc.GetElement(dep_id)
                    if dep_el and is_gtp_element(dep_el) and dep_el not in gtps:
                        gtps.append(dep_el)
        except:
            pass

    return gtps


# ── collect elements ──────────────────────────────────────────────────────────

def collect_elements():
    view_elements = list(
        FilteredElementCollector(doc, curview.Id)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    temp_list = []

    for el in view_elements:
        if not el or not el.Category:
            continue

        try:
            cat_id = get_id_value(el.Category.Id)
            if cat_id == int(BuiltInCategory.OST_FabricationHangers):
                temp_list.append(el)
            else:
                temp_list.extend(get_all_gtps_from_element(el))
        except:
            pass

    seen = set()
    all_elements = []
    for el in temp_list:
        el_id = get_element_id_value(el)
        if el_id in seen:
            continue
        seen.add(el_id)
        all_elements.append(el)

    filtered = [
        el for el in all_elements
        if not (is_fab_hanger(el) and is_beam_hanger(el))
    ]

    return sorted(filtered, key=lambda el: get_element_id_value(el))


# ── build point lines ─────────────────────────────────────────────────────────

def build_point_lines(elements):
    lines = ["SETAPLCOLOR "]

    owner_by_element = {}
    gtp_groups_by_owner = defaultdict(list)

    for element in elements:
        owner = get_point_number_owner(element)
        owner_by_element[get_element_id_value(element)] = owner

        if not is_fab_hanger(element):
            owner_id = get_element_id_value(owner)
            gtp_groups_by_owner[owner_id].append(element)

    gtp_index_by_element = {}
    gtp_count_by_owner = {}

    for owner_id, gtp_list in gtp_groups_by_owner.iteritems():
        sorted_gtps = sorted(gtp_list, key=lambda el: get_element_id_value(el))
        gtp_count_by_owner[owner_id] = len(sorted_gtps)
        for idx, gtp_el in enumerate(sorted_gtps):
            gtp_index_by_element[get_element_id_value(gtp_el)] = idx

    point_count = 0

    for element in elements:
        owner = owner_by_element[get_element_id_value(element)]
        owner_number = get_ts_point_number(owner)

        if not owner_number:
            continue

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
                        pid = owner_number
                    else:
                        pid = make_child_point_number(owner_number, i)

                    north, east, elev = get_shared_coords(pos)
                    if north is None:
                        continue

                    lines.append("POINT {},{},{} ".format(
                        round(east * 12, 8),
                        round(north * 12, 8),
                        round(elev * 12, 8)
                    ))
                    point_count += 1
            except:
                continue

        else:
            try:
                loc = element.Location
                if hasattr(loc, 'Point') and loc.Point is not None:
                    pos = loc.Point
                else:
                    pos = element.GetTransform().Origin

                north, east, elev = get_shared_coords(pos)
                if north is None:
                    continue

                owner_id = get_element_id_value(owner)
                gtp_count = gtp_count_by_owner.get(owner_id, 1)
                gtp_index = gtp_index_by_element.get(get_element_id_value(element), 0)

                if gtp_count == 1:
                    pid = owner_number
                else:
                    pid = make_child_point_number(owner_number, gtp_index)

                lines.append("POINT {},{},{} ".format(
                    round(east * 12, 8),
                    round(north * 12, 8),
                    round(elev * 12, 8)
                ))
                point_count += 1
            except:
                continue

    return lines, point_count


# ── display window ────────────────────────────────────────────────────────────

class PointDisplayWindow(Window):
    def __init__(self, lines, point_count):
        self.Title = "Point Export — Copy to CAD"
        self.Width = 680
        self.Height = 520
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.ResizeMode = 0
        self.Topmost = True

        root = Grid()
        root.Margin = Thickness(10)
        self.Content = root

        root.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength(1, GridUnitType.Star)))
        root.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))

        header = Label()
        header.Content = "{} point(s) — Copy, then paste into CAD command line".format(point_count)
        header.FontWeight = FontWeights.Bold
        header.Margin = Thickness(0, 0, 0, 6)
        Grid.SetRow(header, 0)
        root.Children.Add(header)

        scroll = ScrollViewer()
        Grid.SetRow(scroll, 1)
        root.Children.Add(scroll)

        self.txt = TextBox()
        self.txt.FontFamily = FontFamily("Consolas")
        self.txt.FontSize = 12
        self.txt.IsReadOnly = True
        self.txt.AcceptsReturn = True
        self.txt.VerticalScrollBarVisibility = 0
        self.txt.HorizontalScrollBarVisibility = 0
        self.txt.Text = "\r\n".join(lines)
        scroll.Content = self.txt

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
        btn_copy.Click += self.on_copy
        btn_panel.Children.Add(btn_copy)

    def on_copy(self, sender, args):
        self.txt.SelectAll()
        self.txt.Copy()
        self.Close()


# ── main ──────────────────────────────────────────────────────────────────────

elements = collect_elements()

if not elements:
    TaskDialog.Show("No Elements", "No fabrication hangers or GTP points found in the active view.")
else:
    lines, point_count = build_point_lines(elements)

    if point_count == 0:
        TaskDialog.Show("No Points", "Elements were found but none had a TS_Point_Number assigned.")
    else:
        win = PointDisplayWindow(lines, point_count)
        win.ShowDialog()