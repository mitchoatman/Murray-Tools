# -*- coding: UTF-8 -*-
import clr
clr.AddReference('System')
import System
import System.Diagnostics
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import Window, Thickness, HorizontalAlignment, WindowStartupLocation, ResizeMode
from System.Windows.Controls import Grid, RowDefinition, Label, TextBox, Button, StackPanel, Orientation, RadioButton
from System.Windows.Interop import WindowInteropHelper

from Autodesk.Revit import DB
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, Family, ViewType
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Events import TaskDialogShowingEventArgs
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

import os
import sys
import math

# Check active view type
view = __revit__.ActiveUIDocument.ActiveView
if view.ViewType == ViewType.ThreeD:
    TaskDialog.Show("Error", "Cannot use in 3D view.")
    sys.exit()

path, filename = os.path.split(__file__)
family_file = '\\Control Point.rfa'

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
uiapp = __revit__

FamilyName = 'Control Point'
TypeName = 'CP'

# Save file
folder_name = "c:\\temp"
filepath = os.path.join(folder_name, 'Ribbon_Control Point.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)

default_point_number = ""
default_point_name = ""
default_mode = "Append"

if os.path.exists(filepath):
    try:
        with open(filepath, 'r') as f:
            lines = f.read().splitlines()
            if len(lines) > 0:
                default_point_number = lines[0].strip()
            if len(lines) > 1:
                default_point_name = lines[1].strip()
            if len(lines) > 2 and lines[2].strip() in ["Append", "Custom"]:
                default_mode = lines[2].strip()
    except:
        pass


class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Project
        overwriteParameterValues.Value = False
        return True


def get_family(document, family_name):
    families = FilteredElementCollector(document).OfClass(Family)
    for fam in families:
        if fam.Name == family_name:
            return fam
    return None


def get_symbol_name(symbol):
    try:
        return symbol.Name
    except:
        pass

    try:
        param = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if param:
            return param.AsString()
    except:
        pass

    return None


def get_family_symbol_by_name(document, family_name, type_name):
    family = get_family(document, family_name)
    if not family:
        return None

    for symbol_id in family.GetFamilySymbolIds():
        symbol = document.GetElement(symbol_id)
        if symbol:
            symbol_name = get_symbol_name(symbol)
            if symbol_name and symbol_name.upper() == type_name.upper():
                return symbol

    return None


def set_text_parameter(element, param_name, value):
    param = element.LookupParameter(param_name)
    if not param:
        return "Parameter '{}' not found.".format(param_name)
    if param.IsReadOnly:
        return "Parameter '{}' is read-only.".format(param_name)
    if param.StorageType != DB.StorageType.String:
        return "Parameter '{}' is not a text parameter.".format(param_name)

    param.Set(value)
    return None


def shared_family_dialog_fallback(sender, args):
    try:
        if isinstance(args, TaskDialogShowingEventArgs):
            msg = (args.Message or "").lower()
            dialog_id = (args.DialogId or "").lower()

            if ("shared" in msg and "already exists" in msg and "project" in msg) \
               or ("shared" in dialog_id and "family" in dialog_id):
                args.OverrideResult(1003)
    except:
        pass


class GridSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return isinstance(elem, DB.Grid)

    def AllowReference(self, reference, position):
        return False


def select_grids():
    refs = uidoc.Selection.PickObjects(
        ObjectType.Element,
        GridSelectionFilter(),
        "Select 2 or more grid lines"
    )

    grids = []
    for r in refs:
        elem = doc.GetElement(r.ElementId)
        if isinstance(elem, DB.Grid):
            grids.append(elem)

    if len(grids) < 2:
        raise Exception("Please select at least 2 grid lines.")

    return grids


def get_grid_direction_xy(grid):
    curve = grid.Curve
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)

    dx = p1.X - p0.X
    dy = p1.Y - p0.Y

    length = math.sqrt((dx * dx) + (dy * dy))
    if length == 0:
        return 0.0, 0.0

    return dx / length, dy / length


def is_more_vertical(grid):
    dx, dy = get_grid_direction_xy(grid)
    return abs(dy) >= abs(dx)


def get_y_and_x_grids(grid1, grid2):
    g1_vertical = is_more_vertical(grid1)
    g2_vertical = is_more_vertical(grid2)

    if g1_vertical and not g2_vertical:
        return grid1, grid2

    if g2_vertical and not g1_vertical:
        return grid2, grid1

    dx1, dy1 = get_grid_direction_xy(grid1)
    dx2, dy2 = get_grid_direction_xy(grid2)

    score1 = abs(dy1) - abs(dx1)
    score2 = abs(dy2) - abs(dx2)

    if score1 >= score2:
        return grid1, grid2
    else:
        return grid2, grid1


def get_curve_intersection_points(curve1, curve2):
    points = []

    # Revit 2026/2027+ path
    if hasattr(DB, "CurveIntersectResultOption"):
        try:
            inter = curve1.Intersect(curve2, DB.CurveIntersectResultOption.Detailed)
            if inter and inter.Result == DB.SetComparisonResult.Overlap:
                for ov in inter.GetOverlaps():
                    p = ov.Point
                    if p:
                        points.append(p)
            return points
        except:
            pass

    # Legacy path
    try:
        results_ref = clr.Reference[DB.IntersectionResultArray]()
        result = curve1.Intersect(curve2, results_ref)

        if result == DB.SetComparisonResult.Overlap and results_ref.Value:
            for k in range(results_ref.Value.Size):
                p = results_ref.Value.get_Item(k).XYZPoint
                if p:
                    points.append(p)
    except:
        pass

    return points


def get_grid_intersection_data(grids):
    data = []
    seen = set()

    for i in range(len(grids)):
        curve1 = grids[i].Curve

        for j in range(i + 1, len(grids)):
            curve2 = grids[j].Curve

            points = get_curve_intersection_points(curve1, curve2)
            if not points:
                continue

            y_grid, x_grid = get_y_and_x_grids(grids[i], grids[j])

            for p in points:
                key = (round(p.X, 6), round(p.Y, 6), round(p.Z, 6))
                if key not in seen:
                    seen.add(key)
                    data.append({
                        "point": p,
                        "y_grid_name": y_grid.Name,
                        "x_grid_name": x_grid.Name
                    })

    return data


def get_view_level(view_obj):
    try:
        if view_obj.GenLevel:
            return view_obj.GenLevel
    except:
        pass

    try:
        if view_obj.LevelId != DB.ElementId.InvalidElementId:
            return doc.GetElement(view_obj.LevelId)
    except:
        pass

    return None


def create_control_point_instance(point, symbol, view_obj):
    placement_type = symbol.Family.FamilyPlacementType

    if placement_type == DB.FamilyPlacementType.ViewBased:
        return doc.Create.NewFamilyInstance(point, symbol, view_obj)

    level = get_view_level(view_obj)
    if level:
        try:
            return doc.Create.NewFamilyInstance(
                point,
                symbol,
                level,
                DB.Structure.StructuralType.NonStructural
            )
        except:
            pass

    return doc.Create.NewFamilyInstance(
        point,
        symbol,
        DB.Structure.StructuralType.NonStructural
    )


def build_point_number(mode, custom_value, y_grid_name, x_grid_name):
    if mode == "Custom":
        return custom_value
    return "CP_{}-{}".format(y_grid_name, x_grid_name)


class ControlPointWindow(Window):
    def __init__(self, revit_window_handle, default_number, default_name, default_mode):
        self.Title = "Set Control Point Data"
        self.Width = 320
        self.Height = 255
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen

        self.point_number = default_number
        self.point_name = default_name
        self.naming_mode = default_mode
        self.confirmed = False

        self.InitializeComponents()
        WindowInteropHelper(self).Owner = revit_window_handle

    def InitializeComponents(self):
        grid = Grid()
        self.Content = grid

        row_definitions = [
            RowDefinition(Height=System.Windows.GridLength.Auto),
            RowDefinition(Height=System.Windows.GridLength.Auto),
            RowDefinition(Height=System.Windows.GridLength.Auto),
            RowDefinition(Height=System.Windows.GridLength.Auto),
            RowDefinition(Height=System.Windows.GridLength.Auto),
            RowDefinition(Height=System.Windows.GridLength.Auto),
            RowDefinition(Height=System.Windows.GridLength.Auto)
        ]
        for row in row_definitions:
            grid.RowDefinitions.Add(row)

        row_index = 0

        label_mode = Label()
        label_mode.Content = "Point Number Mode:"
        label_mode.Margin = Thickness(10, 5, 10, 2)
        Grid.SetRow(label_mode, row_index)
        grid.Children.Add(label_mode)
        row_index += 1

        radio_panel = StackPanel()
        radio_panel.Orientation = Orientation.Horizontal
        radio_panel.Margin = Thickness(10, 0, 10, 5)
        radio_panel.HorizontalAlignment = HorizontalAlignment.Left
        Grid.SetRow(radio_panel, row_index)
        grid.Children.Add(radio_panel)

        self.radio_append = RadioButton()
        self.radio_append.Content = "Append Grid Names"
        self.radio_append.Margin = Thickness(0, 0, 20, 0)
        self.radio_append.Checked += self.on_mode_checked
        radio_panel.Children.Add(self.radio_append)

        self.radio_custom = RadioButton()
        self.radio_custom.Content = "Custom"
        self.radio_custom.Margin = Thickness(0, 0, 0, 0)
        self.radio_custom.Checked += self.on_mode_checked
        radio_panel.Children.Add(self.radio_custom)

        row_index += 1

        self.label_number = Label()
        self.label_number.Content = "Point Number:"
        self.label_number.Margin = Thickness(10, 5, 10, 5)
        Grid.SetRow(self.label_number, row_index)
        grid.Children.Add(self.label_number)
        row_index += 1

        self.textbox_number = TextBox()
        self.textbox_number.Text = self.point_number
        self.textbox_number.Margin = Thickness(10, 0, 10, 5)
        Grid.SetRow(self.textbox_number, row_index)
        grid.Children.Add(self.textbox_number)
        row_index += 1

        if self.naming_mode == "Custom":
            self.radio_custom.IsChecked = True
        else:
            self.radio_append.IsChecked = True

        self.label_name = Label()
        self.label_name.Content = "Point Description:"
        self.label_name.Margin = Thickness(10, 5, 10, 5)
        Grid.SetRow(self.label_name, row_index)
        grid.Children.Add(self.label_name)
        row_index += 1

        self.textbox_name = TextBox()
        self.textbox_name.Text = self.point_name
        self.textbox_name.Margin = Thickness(10, 0, 10, 5)
        Grid.SetRow(self.textbox_name, row_index)
        grid.Children.Add(self.textbox_name)
        row_index += 1

        button_panel = StackPanel()
        button_panel.Orientation = Orientation.Horizontal
        button_panel.HorizontalAlignment = HorizontalAlignment.Center
        button_panel.Margin = Thickness(0, 10, 0, 10)
        Grid.SetRow(button_panel, row_index)
        grid.Children.Add(button_panel)

        self.place_button = Button()
        self.place_button.Content = "Place"
        self.place_button.Width = 75
        self.place_button.Height = 25
        self.place_button.Margin = Thickness(5, 0, 5, 0)
        self.place_button.Click += self.on_place_click
        button_panel.Children.Add(self.place_button)

        self.close_button = Button()
        self.close_button.Content = "Close"
        self.close_button.Width = 75
        self.close_button.Height = 25
        self.close_button.Margin = Thickness(5, 0, 5, 0)
        self.close_button.Click += self.on_close_click
        button_panel.Children.Add(self.close_button)

        self.update_point_number_state()

    def update_point_number_state(self):
        if not hasattr(self, 'textbox_number'):
            return

        is_custom = self.radio_custom.IsChecked == True
        self.textbox_number.IsEnabled = is_custom

    def on_mode_checked(self, sender, event):
        self.update_point_number_state()

    def on_place_click(self, sender, event):
        self.point_number = self.textbox_number.Text.strip()
        self.point_name = self.textbox_name.Text.strip()
        self.naming_mode = "Custom" if self.radio_custom.IsChecked else "Append"

        if self.naming_mode == "Custom" and not self.point_number:
            TaskDialog.Show("Error", "Please enter a Point Number.")
            return

        if not self.point_name:
            TaskDialog.Show("Error", "Please enter a Point Description.")
            return

        self.confirmed = True
        self.Close()

    def on_close_click(self, sender, event):
        self.confirmed = False
        self.Close()


family_path = path + family_file

family = get_family(doc, FamilyName)

if not family:
    fload_handler = FamilyLoaderOptionsHandler()
    loaded_family_ref = clr.Reference[DB.Family]()

    uiapp.DialogBoxShowing += shared_family_dialog_fallback
    try:
        t = Transaction(doc, 'Load Control Point Family')
        t.Start()
        try:
            doc.LoadFamily(family_path, fload_handler, loaded_family_ref)
            t.Commit()
        except Exception as e:
            if t.HasStarted():
                t.RollBack()
            TaskDialog.Show("Error", "Error loading family: {}".format(str(e)))
            sys.exit()
    finally:
        uiapp.DialogBoxShowing -= shared_family_dialog_fallback

family = get_family(doc, FamilyName)
if not family:
    TaskDialog.Show("Error", "Family '{}' was not found in the project.".format(FamilyName))
    sys.exit()

target_symbol = get_family_symbol_by_name(doc, FamilyName, TypeName)
if not target_symbol:
    TaskDialog.Show("Error", "Type '{}' was not found in family '{}'.".format(TypeName, FamilyName))
    sys.exit()

if not target_symbol.IsActive:
    t = Transaction(doc, 'Activate Control Point Type')
    t.Start()
    try:
        target_symbol.Activate()
        doc.Regenerate()
        t.Commit()
    except Exception as e:
        TaskDialog.Show("Error", "Error activating type '{}': {}".format(TypeName, str(e)))
        if t.HasStarted():
            t.RollBack()
        sys.exit()

try:
    grids = select_grids()
    intersection_data = get_grid_intersection_data(grids)

    if not intersection_data:
        TaskDialog.Show("Error", "No intersections were found between the selected grid lines.")
        sys.exit()

except Exception as e:
    TaskDialog.Show("Error", str(e))
    sys.exit()

revit_window_handle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle
form = ControlPointWindow(revit_window_handle, default_point_number, default_point_name, default_mode)
form.ShowDialog()

if not form.confirmed:
    sys.exit()

try:
    with open(filepath, 'w') as f:
        f.write(form.point_number + '\n')
        f.write(form.point_name + '\n')
        f.write(form.naming_mode)
except:
    pass

t = Transaction(doc, 'Place Control Points at Grid Intersections')
t.Start()
try:
    errors = []
    placed_count = 0

    for item in intersection_data:
        point = item["point"]
        y_grid_name = item["y_grid_name"]
        x_grid_name = item["x_grid_name"]

        point_number_value = build_point_number(
            form.naming_mode,
            form.point_number,
            y_grid_name,
            x_grid_name
        )

        elem = create_control_point_instance(point, target_symbol, view)
        if not elem:
            continue

        err1 = set_text_parameter(elem, "TS_Point_Number", point_number_value)
        err2 = set_text_parameter(elem, "TS_Point_Description", form.point_name)

        if err1:
            errors.append(err1)
        if err2:
            errors.append(err2)

        placed_count += 1

    t.Commit()

    if errors:
        unique_errors = sorted(set(errors))
        TaskDialog.Show(
            "Warning",
            "Placed {} control point(s).\n\n{}".format(placed_count, "\n".join(unique_errors))
        )
    else:
        TaskDialog.Show("Complete", "Placed {} control point(s).".format(placed_count))

except Exception as e:
    if t.HasStarted():
        t.RollBack()
    TaskDialog.Show("Error", "Error placing control points: {}".format(str(e)))