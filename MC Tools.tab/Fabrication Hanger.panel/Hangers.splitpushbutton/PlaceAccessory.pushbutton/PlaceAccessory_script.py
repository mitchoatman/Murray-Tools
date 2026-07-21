import Autodesk
from Autodesk.Revit.DB import (
    Transaction,
    BuiltInParameter,
    FamilySymbol,
    XYZ,
    ElementTransformUtils,
    Line,
    FilteredElementCollector,
    BuiltInCategory
)
from Autodesk.Revit.UI.Selection import ObjectType
import math
import os
import re
import clr
from fractions import Fraction

clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from Autodesk.Revit.UI import TaskDialog
from System.Windows import Window, Thickness, WindowStartupLocation, ResizeMode, HorizontalAlignment
from System.Windows.Controls import StackPanel, Label, ComboBox, TextBox, CheckBox, Button, Orientation
from System.Windows.Media import FontFamily
from System import Array

# ------------------------------------------------------------------------------------
# REVIT CONTEXT
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

# ------------------------------------------------------------------------------------
# HELPERS

def frac2string(s):
    i, f = s.groups(0)
    f = Fraction(f)
    return str(int(i) + float(f))

def get_parameter_value(element, parameterName):
    param = element.LookupParameter(parameterName)
    if param and param.HasValue:
        try:
            return param.AsDouble()
        except:
            return None
    return None

def get_curve_from_element(ele):
    try:
        if ele.Location and hasattr(ele.Location, "Curve"):
            return ele.Location.Curve
    except:
        pass
    try:
        return ele.get_Curve()
    except:
        return None

def get_xy_direction_and_endpoints(ele):
    curve = get_curve_from_element(ele)
    if not curve:
        raise Exception("No valid curve found for element {}".format(ele.Id))

    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    vec_xy = XYZ(p1.X - p0.X, p1.Y - p0.Y, 0)

    if vec_xy.GetLength() == 0:
        raise Exception("Element {} has no valid XY direction.".format(ele.Id))

    return curve, p0, p1, vec_xy.Normalize()

def get_insulation_thickness(ele):
    try:
        if hasattr(ele, "HasInsulation") and ele.HasInsulation and hasattr(ele, "InsulationThickness"):
            return ele.InsulationThickness
    except:
        pass
    try:
        if hasattr(ele, "InsulationThickness") and ele.InsulationThickness > 0:
            return ele.InsulationThickness
    except:
        pass
    return 0.0

def get_numeric_outside_diameter(ele):
    try:
        d = ele.Diameter
        if d and d > 0:
            return float(d)
    except:
        pass

    for pname in ["Outside Diameter", "Diameter"]:
        p = ele.LookupParameter(pname)
        if p and p.HasValue:
            try:
                return p.AsDouble()
            except:
                pass

    try:
        connectors = list(ele.ConnectorManager.Connectors)
        for conn in connectors:
            try:
                if conn.Radius > 0:
                    return conn.Radius * 2.0
            except:
                pass
    except:
        pass

    return 0.0

def get_bottom_elevation(ele, include_insulation=False):
    bottom = None

    if RevitINT > 2022:
        bottom = get_parameter_value(ele, 'Lower End Bottom Elevation')
    if bottom is None and RevitINT < 2023:
        bottom = get_parameter_value(ele, 'Bottom')

    if bottom is None:
        bbox = ele.get_BoundingBox(None)
        if bbox:
            bottom = bbox.Min.Z

    if bottom is None:
        curve = get_curve_from_element(ele)
        if curve:
            bottom = min(curve.GetEndPoint(0).Z, curve.GetEndPoint(1).Z)

    if bottom is None:
        raise Exception("Cannot determine bottom elevation for element {}".format(ele.Id))

    if include_insulation:
        bottom -= get_insulation_thickness(ele)

    return bottom

def GetCenterPoint(ele_id):
    bBox = doc.GetElement(ele_id).get_BoundingBox(None)
    if bBox is None:
        return XYZ(0, 0, 0)
    center = (bBox.Max + bBox.Min) / 2
    return center

def myround(x, multiple):
    return multiple * math.ceil(x / multiple)

def get_reference_level(hanger):
    level_id = hanger.LevelId
    return doc.GetElement(level_id)

def get_level_elevation(level):
    if level:
        try:
            return level.ProjectElevation
        except:
            return level.Elevation
    return 0.0

# ------------------------------------------------------------------------------------
# DIALOG

class HangerSpacingDialog(Window):
    def __init__(self, family_names, lines, checkboxdefBOI, checkboxdefRotate):
        super(HangerSpacingDialog, self).__init__()
        self.Title = "Hanger and Spacing"
        self.Width = 390
        self.Height = 300
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.NoResize

        stack = StackPanel()
        stack.Orientation = Orientation.Vertical
        stack.Margin = Thickness(10)

        label_hanger = Label()
        label_hanger.Content = "Choose Hanger Family:"
        label_hanger.FontSize = 12
        label_hanger.FontFamily = FontFamily("Arial")
        stack.Children.Add(label_hanger)

        self.combobox_hanger = ComboBox()
        self.combobox_hanger.Width = 350
        self.combobox_hanger.Height = 20
        self.combobox_hanger.FontSize = 12
        self.combobox_hanger.FontFamily = FontFamily("Arial")
        self.combobox_hanger.ItemsSource = Array[object](family_names)
        if lines[0] in family_names:
            self.combobox_hanger.SelectedItem = lines[0]
        else:
            self.combobox_hanger.SelectedItem = family_names[0] if family_names else None
        self.combobox_hanger.Margin = Thickness(0, 0, 0, 10)
        self.combobox_hanger.HorizontalAlignment = HorizontalAlignment.Left
        stack.Children.Add(self.combobox_hanger)

        label_end_dist = Label()
        label_end_dist.Content = "Distance from End (Ft):"
        label_end_dist.FontSize = 12
        label_end_dist.FontFamily = FontFamily("Arial")
        stack.Children.Add(label_end_dist)

        self.textbox_end_dist = TextBox()
        self.textbox_end_dist.Width = 200
        self.textbox_end_dist.Height = 20
        self.textbox_end_dist.FontSize = 12
        self.textbox_end_dist.FontFamily = FontFamily("Arial")
        self.textbox_end_dist.Text = lines[1]
        self.textbox_end_dist.Margin = Thickness(0, 0, 0, 10)
        self.textbox_end_dist.HorizontalAlignment = HorizontalAlignment.Left
        stack.Children.Add(self.textbox_end_dist)

        label_spacing = Label()
        label_spacing.Content = "Hanger Spacing (Ft):"
        label_spacing.FontSize = 12
        label_spacing.FontFamily = FontFamily("Arial")
        stack.Children.Add(label_spacing)

        self.textbox_spacing = TextBox()
        self.textbox_spacing.Width = 200
        self.textbox_spacing.Height = 20
        self.textbox_spacing.FontSize = 12
        self.textbox_spacing.FontFamily = FontFamily("Arial")
        self.textbox_spacing.Text = lines[2]
        self.textbox_spacing.Margin = Thickness(0, 0, 0, 10)
        self.textbox_spacing.HorizontalAlignment = HorizontalAlignment.Left
        stack.Children.Add(self.textbox_spacing)

        self.checkbox_boi = CheckBox()
        self.checkbox_boi.Content = "Align Trapeze to Bottom of Insulation"
        self.checkbox_boi.FontSize = 12
        self.checkbox_boi.FontFamily = FontFamily("Arial")
        self.checkbox_boi.IsChecked = checkboxdefBOI
        self.checkbox_boi.Margin = Thickness(0, 0, 0, 5)
        stack.Children.Add(self.checkbox_boi)

        self.checkbox_rotate = CheckBox()
        self.checkbox_rotate.Content = "Rotate Family"
        self.checkbox_rotate.FontSize = 12
        self.checkbox_rotate.FontFamily = FontFamily("Arial")
        self.checkbox_rotate.IsChecked = checkboxdefRotate
        self.checkbox_rotate.Margin = Thickness(0, 0, 0, 10)
        stack.Children.Add(self.checkbox_rotate)

        self.button_ok = Button()
        self.button_ok.Content = "OK"
        self.button_ok.FontSize = 12
        self.button_ok.FontFamily = FontFamily("Arial")
        self.button_ok.Width = 74
        self.button_ok.Height = 25
        self.button_ok.HorizontalAlignment = HorizontalAlignment.Center
        self.button_ok.Click += self.ok_button_clicked
        stack.Children.Add(self.button_ok)

        self.Content = stack

    def ok_button_clicked(self, sender, event):
        self.DialogResult = True
        self.Close()

# ------------------------------------------------------------------------------------
# MAIN

try:
    selected_reference = uidoc.Selection.PickObject(ObjectType.Element, 'Select OUTSIDE Pipe')
    selected_element = selected_reference
    pick_point = selected_reference.GlobalPoint

    element = doc.GetElement(selected_element.ElementId)
    selected_element1 = uidoc.Selection.PickObject(ObjectType.Element, 'Select OPPOSITE OUTSIDE Pipe')
    element1 = doc.GetElement(selected_element1.ElementId)
    selected_elements = [element, element1]

    level_id = element.LevelId
    level = doc.GetElement(level_id)
    level_elevation = level.Elevation if level else 0

    # Get bottom elevation of selected pipe
    PRTElevation = None
    if element and RevitINT > 2022:
        PRTElevation = get_parameter_value(element, 'Lower End Bottom Elevation')
    if element and RevitINT < 2023:
        PRTElevation = get_parameter_value(element, 'Bottom')

    if PRTElevation is None:
        curve = get_curve_from_element(element)
        if curve:
            PRTElevation = min(curve.GetEndPoint(0).Z, curve.GetEndPoint(1).Z)
        else:
            connectors = element.ConnectorManager.Connectors
            connector_list = list(connectors)
            if connector_list:
                PRTElevation = min([conn.Origin.Z for conn in connector_list])
            else:
                raise Exception("Cannot determine pipe elevation.")

    if PRTElevation < level_elevation - 1.0:
        bbox = element.get_BoundingBox(None)
        if bbox:
            PRTElevation = bbox.Min.Z

    # Get Outside Diameter string of first selected pipe
    outside_diameter = None
    od_param = element.LookupParameter("Overall Size")
    if od_param and od_param.HasValue:
        outside_diameter = od_param.AsString()
    else:
        try:
            outside_diameter = element.Diameter
        except:
            outside_diameter = None

    # Collect all pipe accessory families
    pipe_accessory_symbols = FilteredElementCollector(doc)\
        .OfClass(FamilySymbol)\
        .OfCategory(BuiltInCategory.OST_PipeAccessory)\
        .ToElements()

    family_names = []
    family_symbol_dict = {}

    for symbol in pipe_accessory_symbols:
        family_name = symbol.Family.Name
        type_name = symbol.LookupParameter("Type Name").AsString() if symbol.LookupParameter("Type Name") else symbol.Name
        display_name = family_name + " - " + type_name
        if display_name not in family_names:
            family_names.append(display_name)
            family_symbol_dict[display_name] = symbol

    family_names.sort()

    folder_name = "c:\\Temp"
    filepath = os.path.join(folder_name, 'Ribbon_PlaceAccessory.txt')

    if not os.path.exists(folder_name):
        os.makedirs(folder_name)

    if not os.path.exists(filepath):
        with open(filepath, 'w') as the_file:
            line1 = (family_names[0] + '\n') if family_names else ('Unknown Family - Unknown Type' + '\n')
            line2 = '1.0\n'
            line3 = '8.0\n'
            line4 = 'True\n'
            line5 = 'True\n'
            the_file.writelines([line1, line2, line3, line4, line5])

    with open(filepath, 'r') as file:
        lines = [line.rstrip() for line in file.readlines()]

    if len(lines) < 5:
        with open(filepath, 'w') as the_file:
            line1 = (family_names[0] + '\n') if family_names else ('Unknown Family - Unknown Type' + '\n')
            line2 = '1.0\n'
            line3 = '8.0\n'
            line4 = 'True\n'
            line5 = 'True\n'
            the_file.writelines([line1, line2, line3, line4, line5])

    with open(filepath, 'r') as file:
        lines = [line.rstrip() for line in file.readlines()]

    checkboxdefBOI = False if lines[3] == 'False' else True
    checkboxdefRotate = False if (len(lines) > 4 and lines[4] == 'False') else True

    form = HangerSpacingDialog(family_names, lines, checkboxdefBOI, checkboxdefRotate)
    if form.ShowDialog():
        SelectedFamily = str(form.combobox_hanger.SelectedItem)
        distancefromend = form.textbox_end_dist.Text
        Spacing = form.textbox_spacing.Text
        BOITrap = form.checkbox_boi.IsChecked
        RotateFamily = form.checkbox_rotate.IsChecked

        selected_symbol = family_symbol_dict.get(SelectedFamily)
        if not selected_symbol:
            raise Exception("Selected family '{}' not found.".format(SelectedFamily))

        if not selected_symbol.IsActive:
            t = Transaction(doc, 'Activate Family Symbol')
            t.Start()
            selected_symbol.Activate()
            t.Commit()

        with open(filepath, 'w') as the_file:
            the_file.writelines([
                str(SelectedFamily) + '\n',
                str(distancefromend) + '\n',
                str(Spacing) + '\n',
                str(BOITrap) + '\n',
                str(RotateFamily) + '\n'
            ])

        end_distance_val = float(distancefromend)
        spacing_val = float(Spacing)

        if spacing_val <= 0:
            raise Exception("Spacing must be greater than zero.")
        if end_distance_val < 0:
            raise Exception("Distance from end cannot be negative.")

        # -------------------------------------------------------------------------
        # ANGLED RUN SUPPORT

        first_curve, first_p0, first_p1, direction = get_xy_direction_and_endpoints(element)

        if first_p0.DistanceTo(pick_point) <= first_p1.DistanceTo(pick_point):
            reference_point = first_p0
            opposite_end = first_p1
        else:
            reference_point = first_p1
            opposite_end = first_p0

        run_vec_xy = XYZ(
            opposite_end.X - reference_point.X,
            opposite_end.Y - reference_point.Y,
            0
        )

        run_length = run_vec_xy.GetLength()
        if run_length == 0:
            raise Exception("Run length is zero.")

        perpendicular = XYZ(-direction.Y, direction.X, 0).Normalize()
        rotation_angle = math.atan2(direction.Y, direction.X)

        edge_projections = []
        placement_z = float('inf')

        for pipe in selected_elements:
            curve, p0, p1, pipe_dir = get_xy_direction_and_endpoints(pipe)
            midpoint = (p0 + p1) / 2
            insulation = get_insulation_thickness(pipe)
            outside_dia_num = get_numeric_outside_diameter(pipe)

            half_width = (outside_dia_num / 2.0) + insulation
            perp_offset = (midpoint - reference_point).DotProduct(perpendicular)

            edge_projections.append(perp_offset - half_width)
            edge_projections.append(perp_offset + half_width)

            bottom_z = get_bottom_elevation(pipe, include_insulation=BOITrap)
            placement_z = min(placement_z, bottom_z)

        if not edge_projections:
            raise Exception("Could not determine trapeze width.")

        min_perp = min(edge_projections)
        max_perp = max(edge_projections)
        trapeze_center_offset = (min_perp + max_perp) / 2.0
        trapeze_width = max_perp - min_perp

        if trapeze_width <= 0:
            raise Exception("Calculated trapeze width is invalid.")

        if end_distance_val > run_length:
            raise Exception("Distance from end exceeds run length.")

        qtyofhgrs = int(math.floor((run_length - end_distance_val) / spacing_val)) + 1
        if qtyofhgrs <= 0:
            raise Exception("Calculated hanger quantity is zero.")

        base_width = myround((trapeze_width * 12.0), 2) / 12.0
        is_figure_109 = SelectedFamily == "CEAS Stiffy Figure 109 - 109"
        newwidth = base_width if is_figure_109 else base_width + (4.0 / 12.0)

        TIER_1_CLEARANCE_FT = 0.0
        first_pipe_od = get_numeric_outside_diameter(element)

        hangers = []
        t = Transaction(doc, 'Place Pipe Accessory Hanger')
        t.Start()

        for idx in range(qtyofhgrs):
            along_distance = end_distance_val + (idx * spacing_val)

            xy_point = reference_point + (direction * along_distance) + (perpendicular * trapeze_center_offset)
            location = XYZ(xy_point.X, xy_point.Y, placement_z)

            hanger = doc.Create.NewFamilyInstance(
                location,
                selected_symbol,
                doc.GetElement(level_id),
                Autodesk.Revit.DB.Structure.StructuralType.NonStructural
            )
            hangers.append(hanger)

        doc.Regenerate()

        for hanger in hangers:
            # Set width parameters
            dim_c_param = hanger.LookupParameter("DIM C")
            if dim_c_param:
                dim_c_param.Set(newwidth)

            if is_figure_109:
                dim_b_param = hanger.LookupParameter("DIM B")
                if dim_b_param:
                    dim_b_param.Set(newwidth)

            # Rotate to pipe direction
            center = GetCenterPoint(hanger.Id)
            z_axis_line = Line.CreateBound(center, center + XYZ(0, 0, 1))
            ElementTransformUtils.RotateElement(doc, hanger.Id, z_axis_line, rotation_angle)

            if RotateFamily:
                ElementTransformUtils.RotateElement(doc, hanger.Id, z_axis_line, (90.0 * math.pi / 180.0))

            # Set restraint parameter if available
            if first_pipe_od > 0:
                tier_1_param = hanger.LookupParameter("Tier_1 Restraint")
                if tier_1_param:
                    tier_1_param.Set(first_pipe_od + TIER_1_CLEARANCE_FT)

            # Set offset
            offset_param = hanger.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM)
            if offset_param:
                reference_level = get_reference_level(hanger)
                hanger_level_elevation = get_level_elevation(reference_level)
                offset = placement_z - hanger_level_elevation

                if is_figure_109:
                    dim_a_param = hanger.LookupParameter("DIM A")
                    if dim_a_param and dim_a_param.HasValue:
                        dim_a_value = dim_a_param.AsDouble()
                        offset += dim_a_value - (0.5 / 12.0)

                offset_param.Set(offset)

        t.Commit()

except Exception as e:
    TaskDialog.Show("Error", str(e))