import Autodesk
import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System')
from System.Windows import (
    Window, WindowStartupLocation, WindowStyle, GridLength, HorizontalAlignment,
    VerticalAlignment, Thickness, GridUnitType, TextWrapping, FontWeights
)
from System.Windows.Controls import (
    Label, RadioButton, Button, TextBox, StackPanel, GroupBox, Grid,
    RowDefinition, ColumnDefinition, Orientation, TextBlock, CheckBox
)
from System.Windows.Forms import SaveFileDialog, DialogResult
from System.Diagnostics import Process, ProcessStartInfo
from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    Transaction,
    FamilyInstance
)
from Autodesk.Revit.UI import TaskDialog
import os
import csv
import string
from collections import defaultdict

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)


SETTINGS_FOLDER = r"C:\Temp"
SETTINGS_FILE = os.path.join(SETTINGS_FOLDER, "ExportPoints.txt")


def to_bool(value):
    return str(value).strip().lower() == "true"


def load_export_settings(default_output_path):
    settings = {
        "coordinate_order": "YXZ",
        "coordinate_system": "SHARED",   # default to survey/shared
        "selection_mode": "ALL",
        "use_service_prefix": False,
        "use_item_number": False,
        "output_path": default_output_path
    }

    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, "r") as f:
                for line in f.readlines():
                    line = line.strip()
                    if not line or "=" not in line:
                        continue
                    key, value = line.split("=", 1)
                    key = key.strip()
                    value = value.strip()

                    if key == "coordinate_order":
                        settings["coordinate_order"] = value
                    elif key == "coordinate_system":
                        settings["coordinate_system"] = value
                    elif key == "selection_mode":
                        settings["selection_mode"] = value
                    elif key == "use_service_prefix":
                        settings["use_service_prefix"] = to_bool(value)
                    elif key == "use_item_number":
                        settings["use_item_number"] = to_bool(value)
                    elif key == "output_path":
                        settings["output_path"] = value
    except:
        pass

    return settings


def save_export_settings(settings):
    try:
        if not os.path.exists(SETTINGS_FOLDER):
            os.makedirs(SETTINGS_FOLDER)

        with open(SETTINGS_FILE, "w") as f:
            f.writelines([
                "coordinate_order={}\n".format(settings.get("coordinate_order", "YXZ")),
                "coordinate_system={}\n".format(settings.get("coordinate_system", "SHARED")),
                "selection_mode={}\n".format(settings.get("selection_mode", "ALL")),
                "use_service_prefix={}\n".format(settings.get("use_service_prefix", False)),
                "use_item_number={}\n".format(settings.get("use_item_number", False)),
                "output_path={}\n".format(settings.get("output_path", ""))
            ])
    except:
        pass

def get_id_value(id_obj):
    if id_obj is None:
        return None
    try:
        if RevitINT > 2025:
            return id_obj.Value
        else:
            return id_obj.IntegerValue
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


def make_element_id(id_value):
    try:
        return DB.ElementId(long(id_value))
    except:
        return DB.ElementId(id_value)


def get_suffix(index):
    if index < 26:
        return string.ascii_uppercase[index]
    first = (index // 26) - 1
    second = index % 26
    return string.ascii_uppercase[first] + string.ascii_uppercase[second]


def make_child_point_number(base_id, index):
    return "{}_{}".format(base_id, get_suffix(index))


def is_fab_hanger(element):
    try:
        return element and element.Category and \
               get_id_value(element.Category.Id) == int(BuiltInCategory.OST_FabricationHangers)
    except:
        return False


def normalize_service_abbreviation(value):
    if not value:
        return "UNK"
    cleaned = "".join([c for c in value.strip().upper() if c.isalnum()])
    return cleaned if cleaned else "UNK"


def get_hanger_service_abbreviation(element):
    try:
        abbr = element.ServiceAbbreviation
        if abbr and abbr.strip():
            return normalize_service_abbreviation(abbr)
    except:
        pass

    try:
        p = element.get_Parameter(DB.BuiltInParameter.FABRICATION_SERVICE_ABBREVIATION)
        if p and p.HasValue:
            val = p.AsString()
            if val and val.strip():
                return normalize_service_abbreviation(val)
    except:
        pass

    try:
        p = element.LookupParameter("Fabrication Service Abbreviation")
        if p and p.HasValue:
            val = p.AsString()
            if val and val.strip():
                return normalize_service_abbreviation(val)
    except:
        pass

    try:
        p = element.LookupParameter("Service Abbreviation")
        if p and p.HasValue:
            val = p.AsString()
            if val and val.strip():
                return normalize_service_abbreviation(val)
    except:
        pass

    return "UNK"


def get_parameter_text(element, param_name):
    if not element:
        return ""
    try:
        p = element.LookupParameter(param_name)
        if p:
            val = p.AsString()
            if val and val.strip():
                return val.strip()
            val = p.AsValueString()
            if val and val.strip():
                return val.strip()
    except:
        pass
    return ""


def get_hanger_item_number(element):
    return get_parameter_text(element, "Item Number")


def get_point_number_owner(element):
    try:
        if is_gtp_element(element) and isinstance(element, FamilyInstance) and element.SuperComponent:
            return element.SuperComponent
    except:
        pass
    return element


def get_parent_family_instance(element):
    try:
        if isinstance(element, FamilyInstance) and element.SuperComponent:
            return element.SuperComponent
    except:
        pass
    return None


def clean_z(val, tol=1e-9):
    try:
        return 0.0 if abs(val) < tol else val
    except:
        return val


def parse_point_number(value):
    if not value:
        return ("", None, 1)

    value = value.strip()
    if not value:
        return ("", None, 1)

    i = len(value) - 1
    while i >= 0 and value[i].isdigit():
        i -= 1

    prefix = value[:i+1]
    numeric = value[i+1:]

    if numeric:
        return (prefix, int(numeric), len(numeric))
    else:
        return (value, None, 2)


def make_point_number(prefix, number, width):
    if prefix:
        return prefix + str(number).zfill(width)
    return str(number).zfill(width)


def get_parameter_value_from_parent_first(element, param_name):
    visited = set()
    current = element

    while current:
        try:
            current_id = get_element_id_value(current)
            if current_id in visited:
                break
            visited.add(current_id)
        except:
            pass

        try:
            p = current.LookupParameter(param_name)
            if p and p.HasValue:
                val = p.AsString()
                if val and val.strip():
                    return val.strip()
        except:
            pass

        try:
            if hasattr(current, "Symbol") and current.Symbol:
                p = current.Symbol.LookupParameter(param_name)
                if p and p.HasValue:
                    val = p.AsString()
                    if val and val.strip():
                        return val.strip()
        except:
            pass

        try:
            type_id = current.GetTypeId()
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

        try:
            if isinstance(current, FamilyInstance) and current.SuperComponent:
                current = current.SuperComponent
            else:
                break
        except:
            break

    return None


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


def get_export_coordinates(point, coordinate_system, coord_order):
    if point is None:
        return None, None, None

    if coordinate_system == "SHARED":
        try:
            proj_pos = doc.ActiveProjectLocation.GetProjectPosition(point)
            east = proj_pos.EastWest
            north = proj_pos.NorthSouth
            elev = proj_pos.Elevation
        except:
            east = point.X
            north = point.Y
            elev = point.Z
    else:
        east = point.X
        north = point.Y
        elev = point.Z

    if coord_order == "YXZ":
        return north, east, clean_z(elev)
    else:
        return east, north, clean_z(elev)


class AllElementSelectionFilter(Autodesk.Revit.UI.Selection.ISelectionFilter):
    def AllowElement(self, element):
        return True

    def AllowReference(self, reference, point):
        return False


class ExportHangerPointsDialog(Window):
    def __init__(self):
        self.Title = "Export Hanger Rod & GTP Points to CSV"
        self.Width = 500
        self.Height = 500
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.WindowStyle = WindowStyle.SingleBorderWindow
        self.ResizeMode = 0
        self.Topmost = True

        self.default_filename = "Points.csv"
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        default_full_path = os.path.join(desktop, self.default_filename)

        settings = load_export_settings(default_full_path)

        self.coordinate_order = settings["coordinate_order"]
        self.coordinate_system = settings["coordinate_system"]   # SHARED default
        self.selection_mode = settings["selection_mode"]
        self.use_service_prefix = settings["use_service_prefix"]
        self.use_item_number = settings["use_item_number"]
        self.output_path = settings["output_path"]

        self.InitializeComponents()

    def InitializeComponents(self):
        main_grid = Grid()
        self.Content = main_grid

        for i in range(6):
            row = RowDefinition()
            row.Height = GridLength.Auto
            main_grid.RowDefinitions.Add(row)

        coord_group = GroupBox()
        coord_group.Header = "Point Export Options"
        coord_group.Margin = Thickness(10, 10, 10, 5)
        main_grid.Children.Add(coord_group)
        Grid.SetRow(coord_group, 0)

        coord_grid = Grid()
        coord_grid.Margin = Thickness(10)
        coord_group.Content = coord_grid
        coord_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        coord_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))

        order_panel = StackPanel()
        order_panel.Orientation = Orientation.Vertical
        Grid.SetColumn(order_panel, 0)
        coord_grid.Children.Add(order_panel)

        order_lbl = Label()
        order_lbl.Content = "Axis Order"
        order_lbl.FontWeight = FontWeights.Bold
        order_lbl.Margin = Thickness(0, 0, 0, 5)
        order_panel.Children.Add(order_lbl)

        self.rb_yxz = RadioButton()
        self.rb_yxz.Content = "Y X Z : N.E.H (default)"
        self.rb_yxz.GroupName = "AxisOrder"
        self.rb_yxz.IsChecked = (self.coordinate_order == "YXZ")
        self.rb_yxz.Margin = Thickness(0, 0, 0, 5)
        self.rb_yxz.Checked += lambda s, e: setattr(self, "coordinate_order", "YXZ")
        order_panel.Children.Add(self.rb_yxz)

        self.rb_xyz = RadioButton()
        self.rb_xyz.Content = "X Y Z : E.N.H"
        self.rb_xyz.GroupName = "AxisOrder"
        self.rb_xyz.IsChecked = (self.coordinate_order == "XYZ")
        self.rb_xyz.Margin = Thickness(0, 0, 0, 5)
        self.rb_xyz.Checked += lambda s, e: setattr(self, "coordinate_order", "XYZ")
        order_panel.Children.Add(self.rb_xyz)

        system_panel = StackPanel()
        system_panel.Orientation = Orientation.Vertical
        system_panel.Margin = Thickness(20, 0, 0, 0)
        Grid.SetColumn(system_panel, 1)
        coord_grid.Children.Add(system_panel)

        system_lbl = Label()
        system_lbl.Content = "Coordinate System"
        system_lbl.FontWeight = FontWeights.Bold
        system_lbl.Margin = Thickness(0, 0, 0, 5)
        system_panel.Children.Add(system_lbl)

        self.rb_internal = RadioButton()
        self.rb_internal.Content = "Project Internal"
        self.rb_internal.GroupName = "CoordinateSystem"
        self.rb_internal.IsChecked = (self.coordinate_system == "INTERNAL")
        self.rb_internal.Margin = Thickness(0, 0, 0, 5)
        self.rb_internal.Checked += lambda s, e: setattr(self, "coordinate_system", "INTERNAL")
        system_panel.Children.Add(self.rb_internal)

        self.rb_shared = RadioButton()
        self.rb_shared.Content = "Shared"
        self.rb_shared.GroupName = "CoordinateSystem"
        self.rb_shared.IsChecked = (self.coordinate_system == "SHARED")
        self.rb_shared.Margin = Thickness(0, 0, 0, 5)
        self.rb_shared.Checked += lambda s, e: setattr(self, "coordinate_system", "SHARED")
        system_panel.Children.Add(self.rb_shared)

        select_group = GroupBox()
        select_group.Header = "Point Selection"
        select_group.Margin = Thickness(10, 5, 10, 5)
        main_grid.Children.Add(select_group)
        Grid.SetRow(select_group, 1)

        select_panel = StackPanel()
        select_panel.Orientation = Orientation.Vertical
        select_panel.Margin = Thickness(10)
        select_group.Content = select_panel

        self.rb_all = RadioButton()
        self.rb_all.Content = "All fabrication hangers and GTP points in current view"
        self.rb_all.IsChecked = (self.selection_mode == "ALL")
        self.rb_all.Margin = Thickness(0, 0, 0, 8)
        self.rb_all.Checked += lambda s, e: setattr(self, "selection_mode", "ALL")
        select_panel.Children.Add(self.rb_all)

        self.rb_pick = RadioButton()
        self.rb_pick.Content = "Pick / Window select fabrication hangers and GTP points"
        self.rb_pick.IsChecked = (self.selection_mode == "PICK")
        self.rb_pick.Margin = Thickness(0, 0, 0, 8)
        self.rb_pick.Checked += lambda s, e: setattr(self, "selection_mode", "PICK")
        select_panel.Children.Add(self.rb_pick)

        options_group = GroupBox()
        options_group.Header = "Point Number Options"
        options_group.Margin = Thickness(10, 5, 10, 5)
        main_grid.Children.Add(options_group)
        Grid.SetRow(options_group, 2)

        options_panel = StackPanel()
        options_panel.Orientation = Orientation.Vertical
        options_panel.Margin = Thickness(10)
        options_group.Content = options_panel

        self.cb_service_prefix = CheckBox()
        self.cb_service_prefix.Content = "Prefix point numbers with service abbreviation"
        self.cb_service_prefix.Margin = Thickness(0, 0, 0, 5)
        options_panel.Children.Add(self.cb_service_prefix)

        self.cb_item_number = CheckBox()
        self.cb_item_number.Content = "Use Item Number as Point Number"
        self.cb_item_number.Margin = Thickness(0, 0, 0, 5)
        options_panel.Children.Add(self.cb_item_number)

        self.cb_service_prefix.Checked += self.on_service_prefix_checked
        self.cb_service_prefix.Unchecked += self.on_service_prefix_unchecked
        self.cb_item_number.Checked += self.on_item_number_checked
        self.cb_item_number.Unchecked += self.on_item_number_unchecked

        self.cb_service_prefix.IsChecked = self.use_service_prefix
        self.cb_item_number.IsChecked = self.use_item_number

        if self.use_service_prefix:
            self.cb_item_number.IsEnabled = False
        elif self.use_item_number:
            self.cb_service_prefix.IsEnabled = False

        file_group = GroupBox()
        file_group.Header = "Output File"
        file_group.Margin = Thickness(10, 5, 10, 5)
        main_grid.Children.Add(file_group)
        Grid.SetRow(file_group, 3)

        file_grid = Grid()
        file_grid.Margin = Thickness(10)
        file_group.Content = file_grid
        file_grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        file_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength.Auto))
        file_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength(1, GridUnitType.Star)))
        file_grid.ColumnDefinitions.Add(ColumnDefinition(Width=GridLength.Auto))

        lbl_name = Label()
        lbl_name.Content = "File:"
        lbl_name.VerticalAlignment = VerticalAlignment.Center
        lbl_name.Margin = Thickness(0, 0, 10, 0)
        Grid.SetRow(lbl_name, 0)
        Grid.SetColumn(lbl_name, 0)
        file_grid.Children.Add(lbl_name)

        self.txt_filename = TextBox()
        self.txt_filename.Text = self.output_path
        self.txt_filename.Margin = Thickness(0, 0, 8, 0)
        Grid.SetRow(self.txt_filename, 0)
        Grid.SetColumn(self.txt_filename, 1)
        file_grid.Children.Add(self.txt_filename)

        btn_browse = Button()
        btn_browse.Content = "Browse..."
        btn_browse.Width = 80
        btn_browse.HorizontalAlignment = HorizontalAlignment.Right
        btn_browse.Click += self.on_browse_clicked
        Grid.SetRow(btn_browse, 0)
        Grid.SetColumn(btn_browse, 2)
        file_grid.Children.Add(btn_browse)

        btn_panel = StackPanel()
        btn_panel.Orientation = Orientation.Horizontal
        btn_panel.HorizontalAlignment = HorizontalAlignment.Center
        btn_panel.Margin = Thickness(0, 15, 0, 10)
        main_grid.Children.Add(btn_panel)
        Grid.SetRow(btn_panel, 5)

        btn_export = Button()
        btn_export.Content = "Export"
        btn_export.Width = 100
        btn_export.Height = 32
        btn_export.Margin = Thickness(0, 0, 20, 0)
        btn_export.Click += self.on_export_clicked
        btn_panel.Children.Add(btn_export)

        btn_cancel = Button()
        btn_cancel.Content = "Cancel"
        btn_cancel.Width = 100
        btn_cancel.Height = 32
        btn_cancel.Click += lambda s, e: self.Close()
        btn_panel.Children.Add(btn_cancel)

    def on_service_prefix_checked(self, sender, args):
        self.use_service_prefix = True
        self.use_item_number = False
        self.cb_item_number.IsChecked = False
        self.cb_item_number.IsEnabled = False

    def on_service_prefix_unchecked(self, sender, args):
        self.use_service_prefix = False
        self.cb_item_number.IsEnabled = True

    def on_item_number_checked(self, sender, args):
        self.use_item_number = True
        self.use_service_prefix = False
        self.cb_service_prefix.IsChecked = False
        self.cb_service_prefix.IsEnabled = False

    def on_item_number_unchecked(self, sender, args):
        self.use_item_number = False
        self.cb_service_prefix.IsEnabled = True

    def on_browse_clicked(self, sender, args):
        dialog = SaveFileDialog()
        dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        dialog.Title = "Save hanger rod and GTP points as CSV"
        dialog.FileName = os.path.basename(self.output_path or self.default_filename)
        dialog.InitialDirectory = os.path.dirname(self.output_path or self.txt_filename.Text)
        dialog.DefaultExt = "csv"
        if dialog.ShowDialog() == DialogResult.OK:
            self.output_path = dialog.FileName
            self.txt_filename.Text = self.output_path

    def on_export_clicked(self, sender, args):
        self.output_path = self.txt_filename.Text.strip() if self.txt_filename.Text else self.output_path
        if not self.output_path:
            TaskDialog.Show("Error", "No file path selected.")
            return

        settings = {
            "coordinate_order": self.coordinate_order,
            "coordinate_system": self.coordinate_system,
            "selection_mode": "PICK" if self.rb_pick.IsChecked else "ALL",
            "use_service_prefix": self.use_service_prefix,
            "use_item_number": self.use_item_number,
            "output_path": self.output_path
        }
        save_export_settings(settings)

        self.DialogResult = True
        self.Close()


class SuccessDialog(Window):
    def __init__(self, filepath, count):
        self.filepath = filepath
        self.count = count
        self.Title = "Export Successful"
        self.Width = 420
        self.Height = 200
        self.ResizeMode = 0
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Topmost = True
        self.InitializeComponents()

    def InitializeComponents(self):
        grid = Grid()
        grid.Margin = Thickness(20)
        self.Content = grid
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))
        grid.RowDefinitions.Add(RowDefinition(Height=GridLength.Auto))

        lbl = Label()
        lbl.Content = "Exported " + str(self.count) + " points successfully."
        lbl.FontSize = 14
        lbl.Margin = Thickness(0, 0, 0, 10)
        Grid.SetRow(lbl, 0)
        grid.Children.Add(lbl)

        path_text = TextBlock()
        path_text.Text = self.filepath
        path_text.FontSize = 12
        path_text.Margin = Thickness(0, 0, 0, 20)
        path_text.TextWrapping = TextWrapping.Wrap
        Grid.SetRow(path_text, 1)
        grid.Children.Add(path_text)

        btn_panel = StackPanel()
        btn_panel.Orientation = Orientation.Horizontal
        btn_panel.HorizontalAlignment = HorizontalAlignment.Center

        open_btn = Button()
        open_btn.Content = "Open File"
        open_btn.Width = 120
        open_btn.Height = 36
        open_btn.Margin = Thickness(0, 0, 20, 0)
        open_btn.Click += self.on_open_clicked
        btn_panel.Children.Add(open_btn)

        close_btn = Button()
        close_btn.Content = "Close"
        close_btn.Width = 120
        close_btn.Height = 36
        close_btn.Click += lambda s, e: self.Close()
        btn_panel.Children.Add(close_btn)

        Grid.SetRow(btn_panel, 2)
        grid.Children.Add(btn_panel)

    def on_open_clicked(self, sender, args):
        try:
            psi = ProcessStartInfo(self.filepath)
            psi.UseShellExecute = True
            Process.Start(psi)
        except Exception as ex:
            TaskDialog.Show("Error", "Could not open the file:\n" + str(ex))
        self.Close()


def perform_export():
    dlg = ExportHangerPointsDialog()
    if dlg.ShowDialog() != True:
        return

    selected_elements = []

    if dlg.rb_pick.IsChecked:
        try:
            filter_obj = AllElementSelectionFilter()
            picked_refs = uidoc.Selection.PickObjects(
                Autodesk.Revit.UI.Selection.ObjectType.Element,
                filter_obj,
                "Select any elements containing hangers or GTP points (window selection now supported)"
            )
            raw_elements = [doc.GetElement(r.ElementId) for r in picked_refs]

            temp_list = []
            for el in raw_elements:
                if el and el.Category:
                    if get_id_value(el.Category.Id) == int(BuiltInCategory.OST_FabricationHangers):
                        temp_list.append(el)
                    else:
                        temp_list.extend(get_all_gtps_from_element(el))

            seen = set()
            selected_elements = [el for el in temp_list if get_element_id_value(el) not in seen and not seen.add(get_element_id_value(el))]

        except Exception as pick_ex:
            TaskDialog.Show("Selection Issue", "Selection was cancelled or failed.\n\nDetail: " + str(pick_ex))
            return
    else:
        hangers = FilteredElementCollector(doc, curview.Id)\
            .OfCategory(BuiltInCategory.OST_FabricationHangers)\
            .WhereElementIsNotElementType()\
            .ToElements()

        generics = FilteredElementCollector(doc, curview.Id)\
            .OfCategory(BuiltInCategory.OST_GenericModel)\
            .WhereElementIsNotElementType()\
            .ToElements()
        gtps = [el for el in generics if is_gtp_element(el)]

        selected_elements = list(hangers) + gtps

    if not selected_elements:
        TaskDialog.Show("No elements", "No valid fabrication hangers or GTP points found.")
        return

    if dlg.rb_yxz.IsChecked:
        coord_order = "YXZ"
        x_header = "Y"
        y_header = "X"
        z_header = "Z"
    else:
        coord_order = "XYZ"
        x_header = "X"
        y_header = "Y"
        z_header = "Z"

    coordinate_system = dlg.coordinate_system

    fieldnames = ['POINT NUMBER', x_header, y_header, z_header, 'DESCRIPTION']
    element_list = sorted(selected_elements, key=lambda el: get_element_id_value(el))

    owner_by_element = {}
    owner_sequence = []
    owner_current_values = {}
    value_count = defaultdict(int)
    seen_owner_ids = set()

    for element in element_list:
        owner = get_point_number_owner(element)
        owner_by_element[get_element_id_value(element)] = owner

        owner_id = get_element_id_value(owner)
        if owner_id in seen_owner_ids:
            continue
        seen_owner_ids.add(owner_id)
        owner_sequence.append(owner)

        current_val = ""

        if is_fab_hanger(owner):
            if dlg.use_item_number:
                item_val = get_hanger_item_number(owner)
                if item_val:
                    current_val = item_val
            elif dlg.use_service_prefix:
                current_val = ""
            else:
                ts_param = owner.LookupParameter("TS_Point_Number")
                if ts_param and ts_param.HasValue:
                    raw_val = ts_param.AsString()
                    if raw_val:
                        current_val = raw_val.strip()
        else:
            ts_param = owner.LookupParameter("TS_Point_Number")
            if ts_param and ts_param.HasValue:
                raw_val = ts_param.AsString()
                if raw_val:
                    current_val = raw_val.strip()

        owner_current_values[owner_id] = current_val
        value_count[current_val] += 1

    assigned_numbers = {}
    assigned_descriptions = {}
    rod_data = []
    skipped_count = 0

    gtp_groups_by_owner = defaultdict(list)

    for element in element_list:
        try:
            if element.Category and get_id_value(element.Category.Id) != int(BuiltInCategory.OST_FabricationHangers):
                owner = owner_by_element[get_element_id_value(element)]
                gtp_groups_by_owner[get_element_id_value(owner)].append(element)
        except:
            pass

    for owner_id in gtp_groups_by_owner:
        gtp_groups_by_owner[owner_id] = sorted(
            gtp_groups_by_owner[owner_id],
            key=lambda el: get_element_id_value(el)
        )

    gtp_index_by_element = {}
    gtp_count_by_owner = {}

    for owner_id, gtp_list in gtp_groups_by_owner.iteritems():
        gtp_count_by_owner[owner_id] = len(gtp_list)
        for idx, gtp_el in enumerate(gtp_list):
            gtp_index_by_element[get_element_id_value(gtp_el)] = idx

    def decimal_to_fraction_inches(decimal_inches):
        if not decimal_inches:
            return ""
        rounded = round(decimal_inches * 16.0) / 16.0
        whole = int(rounded)
        frac = rounded - whole
        if frac < 0.0001:
            return str(whole) if whole > 0 else "<1/16"
        num = int(round(frac * 16))
        den = 16
        for d in range(2, 9):
            while num % d == 0 and den % d == 0:
                num = num // d
                den = den // d
        frac_str = "{}/{}".format(num, den) if den != 1 else str(num)
        if whole > 0:
            return "{}-{}".format(whole, frac_str)
        return frac_str

    used_numbers = set()
    duplicate_next = {}

    for owner in owner_sequence:
        owner_id = get_element_id_value(owner)
        current_val = owner_current_values.get(owner_id, "")

        if not current_val:
            continue

        if value_count[current_val] == 1:
            assigned_numbers[owner_id] = current_val
            used_numbers.add(current_val)
            continue

        prefix, start_num, width = parse_point_number(current_val)

        if start_num is None:
            next_num = duplicate_next.get(current_val, 1)
            candidate = make_point_number(prefix, next_num, width)

            while candidate in used_numbers:
                next_num += 1
                candidate = make_point_number(prefix, next_num, width)

            assigned_numbers[owner_id] = candidate
            used_numbers.add(candidate)
            duplicate_next[current_val] = next_num + 1
        else:
            if current_val not in duplicate_next:
                assigned_numbers[owner_id] = current_val
                used_numbers.add(current_val)
                duplicate_next[current_val] = start_num + 1
            else:
                next_num = duplicate_next[current_val]
                candidate = make_point_number(prefix, next_num, width)

                while candidate in used_numbers:
                    next_num += 1
                    candidate = make_point_number(prefix, next_num, width)

                assigned_numbers[owner_id] = candidate
                used_numbers.add(candidate)
                duplicate_next[current_val] = next_num + 1

    blank_counter = 1
    service_counters = defaultdict(int)

    for owner in owner_sequence:
        owner_id = get_element_id_value(owner)
        if owner_id in assigned_numbers:
            continue

        if is_fab_hanger(owner):
            if dlg.use_service_prefix:
                service_abbr = get_hanger_service_abbreviation(owner)
                service_counters[service_abbr] += 1
                candidate = "{}{}".format(service_abbr, str(service_counters[service_abbr]).zfill(3))
                while candidate in used_numbers:
                    service_counters[service_abbr] += 1
                    candidate = "{}{}".format(service_abbr, str(service_counters[service_abbr]).zfill(3))

            elif dlg.use_item_number:
                candidate = str(blank_counter).zfill(3)
                while candidate in used_numbers:
                    blank_counter += 1
                    candidate = str(blank_counter).zfill(3)
                blank_counter += 1

            else:
                candidate = str(blank_counter).zfill(3)
                while candidate in used_numbers:
                    blank_counter += 1
                    candidate = str(blank_counter).zfill(3)
                blank_counter += 1
        else:
            candidate = str(blank_counter).zfill(3)
            while candidate in used_numbers:
                blank_counter += 1
                candidate = str(blank_counter).zfill(3)
            blank_counter += 1

        assigned_numbers[owner_id] = candidate
        used_numbers.add(candidate)

    for element in element_list:
        owner = owner_by_element[get_element_id_value(element)]
        owner_id = get_element_id_value(owner)
        base_id = assigned_numbers[owner_id]
        is_hanger = is_fab_hanger(element)

        if is_hanger:
            rod_diameter_inches = ""
            try:
                ancillaries = element.GetPartAncillaryUsage()
                for anc in ancillaries:
                    if anc.AncillaryWidthOrDiameter > 0:
                        rod_diameter_inches = anc.AncillaryWidthOrDiameter * 12.0
                        break
            except:
                pass

            frac_dia = decimal_to_fraction_inches(rod_diameter_inches)
            description = "INSERT " + frac_dia if frac_dia else "INSERT"
            assigned_descriptions[get_element_id_value(element)] = description

            try:
                rod_info = element.GetRodInfo()
                if rod_info is None:
                    skipped_count += 1
                    continue
            except:
                skipped_count += 1
                continue

            rod_count = rod_info.RodCount
            for i in range(rod_count):
                try:
                    pos = rod_info.GetRodEndPosition(i)
                    if pos is None:
                        continue

                    point_id = base_id if rod_count == 1 else make_child_point_number(base_id, i)
                    x_val, y_val, z_val = get_export_coordinates(pos, coordinate_system, coord_order)

                    rod_data.append({
                        'POINT NUMBER': point_id,
                        x_header: x_val,
                        y_header: y_val,
                        z_header: z_val,
                        'DESCRIPTION': description
                    })
                except:
                    continue
        else:
            description = "GTP"
            try:
                desc_value = get_parameter_value(element, "TS_Point_Description")
                if isinstance(element, FamilyInstance) and element.SuperComponent:
                    parent_desc = get_parameter_value(element.SuperComponent, "TS_Point_Description")
                    if parent_desc:
                        desc_value = parent_desc
                if desc_value:
                    description = desc_value
            except:
                pass

            try:
                loc = element.Location
                if hasattr(loc, 'Point') and loc.Point is not None:
                    pos = loc.Point
                else:
                    pos = element.GetTransform().Origin

                x_val, y_val, z_val = get_export_coordinates(pos, coordinate_system, coord_order)

                gtp_count = gtp_count_by_owner.get(owner_id, 1)
                gtp_index = gtp_index_by_element.get(get_element_id_value(element), 0)
                point_id = base_id if gtp_count == 1 else make_child_point_number(base_id, gtp_index)

                rod_data.append({
                    'POINT NUMBER': point_id,
                    x_header: x_val,
                    y_header: y_val,
                    z_header: z_val,
                    'DESCRIPTION': description
                })
            except:
                skipped_count += 1
                continue

    if not rod_data:
        TaskDialog.Show("Export Result", "No valid points found.")
        return

    if skipped_count > 0:
        TaskDialog.Show("Warning", "{} element(s) were skipped (no valid position data).".format(skipped_count))

    try:
        with open(dlg.output_path, 'wb') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            for row in rod_data:
                row_str = {k: str(v) for k, v in row.iteritems()}
                writer.writerow(row_str)

        t = Transaction(doc, "Write TS_Point_Number and TS_Point_Description after export")
        t.Start()

        if dlg.use_service_prefix or dlg.use_item_number:
            for owner in owner_sequence:
                if not is_fab_hanger(owner):
                    continue
                try:
                    ts_param = owner.LookupParameter("TS_Point_Number")
                    if ts_param and not ts_param.IsReadOnly:
                        ts_param.Set("")
                except:
                    pass

        for elem_id_int, point_number in assigned_numbers.items():
            target = doc.GetElement(make_element_id(elem_id_int))
            if not target:
                continue
            try:
                ts_param = target.LookupParameter("TS_Point_Number")
                if ts_param and not ts_param.IsReadOnly:
                    ts_param.Set(point_number)
            except:
                pass

        for element in element_list:
            try:
                if is_fab_hanger(element):
                    desc_param = element.LookupParameter("TS_Point_Description")
                    if desc_param and not desc_param.IsReadOnly:
                        desc = assigned_descriptions.get(get_element_id_value(element), "")
                        if desc:
                            desc_param.Set(desc)
            except:
                pass

        t.Commit()

        success = SuccessDialog(dlg.output_path, len(rod_data))
        success.ShowDialog()

    except Exception as ex:
        TaskDialog.Show("Export Failed", str(ex))


perform_export()