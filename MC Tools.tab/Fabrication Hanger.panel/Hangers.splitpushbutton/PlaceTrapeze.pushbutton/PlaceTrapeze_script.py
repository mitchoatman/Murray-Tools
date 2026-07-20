import Autodesk
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.DB import (
    Transaction,
    FabricationConfiguration,
    BuiltInParameter,
    BuiltInCategory,
    FabricationPart,
    XYZ,
    ElementTransformUtils,
    Line,
    ElementId
)
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import math
import os
import clr

clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

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
RevitINT = int(app.VersionNumber)


# ------------------------------------------------------------------------------------
# HELPERS

def get_id_value(elem_id):
    try:
        return int(elem_id.Value)  # Revit 2024+
    except:
        return int(elem_id.IntegerValue)  # Revit 2022/2023


def get_center_point(element_id):
    bbox = doc.GetElement(element_id).get_BoundingBox(None)
    if not bbox:
        raise Exception("No bounding box found for element {}".format(element_id))
    return (bbox.Max + bbox.Min) / 2


def round_up_to_multiple(value, multiple):
    return multiple * math.ceil(value / multiple)


def get_outside_diameter(part):
    param = part.LookupParameter("Outside Diameter")
    if param and param.HasValue:
        return param.AsDouble()
    return 0.0


def get_level_elevation_from_element(element):
    level = doc.GetElement(element.LevelId)
    if not level:
        return 0.0
    try:
        return level.ProjectElevation
    except:
        return level.Elevation


def get_service_names(loaded_services):
    names = []
    for service in loaded_services:
        try:
            names.append(service.Name)
        except:
            names.append([])
    return names


def get_all_hanger_button_names(loaded_services):
    button_names = []
    unique_names = set()

    for service in loaded_services:
        palette_count = service.PaletteCount if RevitINT >= 2023 else service.GroupCount
        for palette_idx in range(palette_count):
            button_count = service.GetButtonCount(palette_idx)
            for btn_idx in range(button_count):
                button = service.GetButton(palette_idx, btn_idx)
                if button.IsAHanger and button.Name not in unique_names:
                    unique_names.add(button.Name)
                    button_names.append(button.Name)

    return button_names


def find_hanger_button(loaded_services, selected_service_name, selected_button_name):
    for service in loaded_services:
        if service.Name != selected_service_name:
            continue

        palette_count = service.PaletteCount if RevitINT >= 2023 else service.GroupCount
        for palette_idx in range(palette_count):
            button_count = service.GetButtonCount(palette_idx)
            for btn_idx in range(button_count):
                button = service.GetButton(palette_idx, btn_idx)
                if button.Name == selected_button_name:
                    return button

    return None


def ensure_settings_file(filepath):
    default_lines = [
        '1.625 Single Strut Trapeze\n',
        '1.0\n',
        '8.0\n',
        'PLUMBING: DOMESTIC COLD WATER\n',
        'True\n',
        'True'
    ]

    folder_name = os.path.dirname(filepath)
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)

    if not os.path.exists(filepath):
        with open(filepath, 'w') as file_obj:
            file_obj.writelines(default_lines)

    with open(filepath, 'r') as file_obj:
        lines = [line.rstrip() for line in file_obj.readlines()]

    if len(lines) < 6:
        with open(filepath, 'w') as file_obj:
            file_obj.writelines(default_lines)
        lines = [line.rstrip() for line in default_lines]

    return lines


def write_settings_file(filepath, hanger_name, end_distance, spacing, service_name, attach_to_structure, boi_value):
    with open(filepath, 'w') as file_obj:
        file_obj.writelines([
            str(hanger_name) + '\n',
            str(end_distance) + '\n',
            str(spacing) + '\n',
            str(service_name) + '\n',
            str(attach_to_structure) + '\n',
            str(boi_value) + '\n'
        ])


# ------------------------------------------------------------------------------------
# SELECTION FILTER

class FabPipeDuctSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        try:
            if elem is None or elem.Category is None or elem.Category.Id is None:
                return False

            cat_id = get_id_value(elem.Category.Id)

            allowed_categories = [
                get_id_value(ElementId(BuiltInCategory.OST_FabricationPipework)),
                get_id_value(ElementId(BuiltInCategory.OST_FabricationDuctwork))
            ]

            return cat_id in allowed_categories
        except:
            return False

    def AllowReference(self, reference, point):
        return True


# ------------------------------------------------------------------------------------
# DIALOG

class HangerSpacingDialog(Window):
    def __init__(self, button_names, service_names, settings_lines, checkboxdefBOI, checkboxdef, is_ptrap=False):
        super(HangerSpacingDialog, self).__init__()

        self.Title = "Hanger and Spacing" if not is_ptrap else "Hanger for P-Trap"
        self.Width = 336
        self.Height = 350 if not is_ptrap else 275
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.NoResize

        stack = StackPanel()
        stack.Orientation = Orientation.Vertical
        stack.Margin = Thickness(10)

        label_hanger = Label()
        label_hanger.Content = "Choose Hanger:"
        label_hanger.FontSize = 12
        label_hanger.FontFamily = FontFamily("Arial")
        stack.Children.Add(label_hanger)

        self.combobox_hanger = ComboBox()
        self.combobox_hanger.Width = 300
        self.combobox_hanger.Height = 20
        self.combobox_hanger.FontSize = 12
        self.combobox_hanger.FontFamily = FontFamily("Arial")
        self.combobox_hanger.ItemsSource = Array[object](button_names)
        if settings_lines[0] in button_names:
            self.combobox_hanger.SelectedItem = settings_lines[0]
        self.combobox_hanger.Margin = Thickness(0, 0, 0, 10)
        self.combobox_hanger.HorizontalAlignment = HorizontalAlignment.Left
        stack.Children.Add(self.combobox_hanger)

        if not is_ptrap:
            label_end_dist = Label()
            label_end_dist.Content = "Distance from End (In):"
            label_end_dist.FontSize = 12
            label_end_dist.FontFamily = FontFamily("Arial")
            stack.Children.Add(label_end_dist)

            self.textbox_end_dist = TextBox()
            self.textbox_end_dist.Width = 200
            self.textbox_end_dist.Height = 20
            self.textbox_end_dist.FontSize = 12
            self.textbox_end_dist.FontFamily = FontFamily("Arial")
            self.textbox_end_dist.Text = str(round(float(settings_lines[1]) * 12.0, 4))
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
            self.textbox_spacing.Text = settings_lines[2]
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

        self.checkbox_attach = CheckBox()
        self.checkbox_attach.Content = "Attach to Structure"
        self.checkbox_attach.FontSize = 12
        self.checkbox_attach.FontFamily = FontFamily("Arial")
        self.checkbox_attach.IsChecked = checkboxdef
        self.checkbox_attach.Margin = Thickness(0, 0, 0, 5 if is_ptrap else 10)
        stack.Children.Add(self.checkbox_attach)

        if is_ptrap:
            label_trap_width = Label()
            label_trap_width.Content = "Trapeze Rod - Rod Width (Ft):"
            label_trap_width.FontSize = 12
            label_trap_width.FontFamily = FontFamily("Arial")
            stack.Children.Add(label_trap_width)

            self.textbox_trap_width = TextBox()
            self.textbox_trap_width.Width = 200
            self.textbox_trap_width.Height = 20
            self.textbox_trap_width.FontSize = 12
            self.textbox_trap_width.FontFamily = FontFamily("Arial")
            self.textbox_trap_width.Text = "1.0"
            self.textbox_trap_width.Margin = Thickness(0, 0, 0, 5)
            self.textbox_trap_width.HorizontalAlignment = HorizontalAlignment.Left
            stack.Children.Add(self.textbox_trap_width)

        label_service = Label()
        label_service.Content = "Choose Service to Draw Hanger on:"
        label_service.FontSize = 12
        label_service.FontFamily = FontFamily("Arial")
        stack.Children.Add(label_service)

        self.combobox_service = ComboBox()
        self.combobox_service.Width = 300
        self.combobox_service.Height = 20
        self.combobox_service.FontSize = 12
        self.combobox_service.FontFamily = FontFamily("Arial")
        self.combobox_service.ItemsSource = Array[object](service_names)
        if settings_lines[3] in service_names:
            self.combobox_service.SelectedItem = settings_lines[3]
        self.combobox_service.Margin = Thickness(0, 0, 0, 10)
        self.combobox_service.HorizontalAlignment = HorizontalAlignment.Left
        stack.Children.Add(self.combobox_service)

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
    selected_reference = uidoc.Selection.PickObject(
        ObjectType.Element,
        FabPipeDuctSelectionFilter(),
        'Select OUTSIDE Pipe'
    )
    element = doc.GetElement(selected_reference.ElementId)
    pick_point = selected_reference.GlobalPoint

    service_param = element.LookupParameter('Fabrication Service')
    if service_param and service_param.HasValue:
        selected_element_service_name = service_param.AsValueString()
    else:
        raise Exception("Fabrication Service parameter missing.")

    config = FabricationConfiguration.GetFabricationConfiguration(doc)
    loaded_services = config.GetAllLoadedServices()
    service_names = get_service_names(loaded_services)

    try:
        service_names.index(selected_element_service_name)
    except ValueError:
        raise Exception("Selected service not found.")

    hanger_button_names = get_all_hanger_button_names(loaded_services)

    settings_path = os.path.join("c:\\Temp", "Ribbon_PlaceTrapeze.txt")
    settings_lines = ensure_settings_file(settings_path)

    checkbox_attach_default = settings_lines[4] != 'False'
    checkbox_boi_default = settings_lines[5] != 'False'

    # --------------------------------------------------------------------------------
    # REGULAR TRAPEZE
    if element.ItemCustomId != 916:
        opposite_reference = uidoc.Selection.PickObject(
            ObjectType.Element,
            FabPipeDuctSelectionFilter(),
            'Select OPPOSITE OUTSIDE Pipe'
        )
        opposite_element = doc.GetElement(opposite_reference.ElementId)
        selected_elements = [element, opposite_element]
        level_id = element.LevelId

        dialog = HangerSpacingDialog(
            hanger_button_names,
            service_names,
            settings_lines,
            checkbox_boi_default,
            checkbox_attach_default,
            is_ptrap=False
        )

        if dialog.ShowDialog():
            selected_hanger_name = str(dialog.combobox_hanger.SelectedItem)
            end_distance_text = dialog.textbox_end_dist.Text
            spacing_text = dialog.textbox_spacing.Text
            align_to_bottom_of_insulation = dialog.checkbox_boi.IsChecked
            attach_to_structure = dialog.checkbox_attach.IsChecked
            selected_service_name = str(dialog.combobox_service.SelectedItem)

            try:
                end_distance = float(end_distance_text) / 12.0
                spacing = float(spacing_text)
            except ValueError:
                raise Exception("Invalid numeric input: Distance from End or Spacing must be numeric.")

            hanger_button = find_hanger_button(loaded_services, selected_service_name, selected_hanger_name)
            if not hanger_button:
                raise Exception("Hanger button '{}' not found in '{}'".format(selected_hanger_name, selected_service_name))

            write_settings_file(
                settings_path,
                selected_hanger_name,
                end_distance,
                spacing,
                selected_service_name,
                attach_to_structure,
                align_to_bottom_of_insulation
            )

            pipe_data = []
            combined_min_z = float('inf')

            for part in selected_elements:
                curve = part.Location.Curve
                if not curve or not curve.IsBound:
                    raise Exception("Pipe {} must have a valid location curve.".format(part.Id))

                point0 = curve.GetEndPoint(0)
                point1 = curve.GetEndPoint(1)
                midpoint = (point0 + point1) / 2
                insulation_thickness = part.InsulationThickness if hasattr(part, 'InsulationThickness') and part.HasInsulation else 0.0
                outside_diameter = get_outside_diameter(part)

                bbox = part.get_BoundingBox(None)
                if not bbox:
                    raise Exception("No bounding box found for pipe {}".format(part.Id))

                bottom_z = bbox.Min.Z
                if align_to_bottom_of_insulation:
                    bottom_z -= insulation_thickness

                combined_min_z = min(combined_min_z, bottom_z)

                pipe_data.append({
                    "element": part,
                    "point0": point0,
                    "point1": point1,
                    "midpoint": midpoint,
                    "insulation_thickness": insulation_thickness,
                    "outside_diameter": outside_diameter
                })

            placement_z = combined_min_z

            first_pipe = pipe_data[0]
            end0 = first_pipe["point0"]
            end1 = first_pipe["point1"]

            if end0.DistanceTo(pick_point) <= end1.DistanceTo(pick_point):
                reference_point = end0
                opposite_end = end1
            else:
                reference_point = end1
                opposite_end = end0

            direction_raw = (opposite_end - reference_point).Normalize()
            direction = XYZ(direction_raw.X, direction_raw.Y, 0)
            if direction.GetLength() == 0:
                raise Exception("Pipe direction could not be determined in XY plane.")
            direction = direction.Normalize()

            perpendicular = XYZ(-direction.Y, direction.X, 0).Normalize()

            edge_projections = []
            for data in pipe_data:
                proj = (data["midpoint"] - reference_point).DotProduct(perpendicular)
                half_size = (data["outside_diameter"] / 2.0) + data["insulation_thickness"]
                edge_projections.append(proj - half_size)
                edge_projections.append(proj + half_size)

            min_perp = min(edge_projections)
            max_perp = max(edge_projections)
            trapeze_center_offset = (min_perp + max_perp) / 2.0
            trapeze_width = max_perp - min_perp

            first_pipe_endpoints = [first_pipe["point0"], first_pipe["point1"]]
            along_projections = [(pt - reference_point).DotProduct(direction) for pt in first_pipe_endpoints]
            min_along = min(along_projections)
            max_along = max(along_projections)
            pipe_length_along = max_along - min_along

            hanger_count = int(math.ceil(pipe_length_along / spacing))
            if hanger_count <= 0:
                raise Exception("Calculated hanger quantity is zero.")

            width_to_set = round_up_to_multiple(trapeze_width * 12.0, 2) / 12.0
            rotation_angle = math.atan2(direction.Y, direction.X)

            transaction = Transaction(doc, 'Place and Modify Trapeze Hanger')
            transaction.Start()

            created_count = 0
            current_spacing = end_distance

            try:
                for index in range(hanger_count):
                    hanger = FabricationPart.CreateHanger(doc, hanger_button, 0, level_id)
                    if not hanger:
                        print("Failed to create hanger {}".format(index + 1))
                        continue

                    created_count += 1

                    for dim in hanger.GetDimensions():
                        dim_name = dim.Name
                        try:
                            if dim_name in ("Width", "Duct Width"):
                                hanger.SetDimensionValue(dim, width_to_set)
                            elif dim_name == "Bearer Extn":
                                hanger.SetDimensionValue(dim, 0.25)
                        except Exception as ex:
                            print("Error setting dimension '{}' for hanger {}: {}".format(dim_name, hanger.Id, str(ex)))

                    doc.Regenerate()

                    center = get_center_point(hanger.Id)
                    z_axis = Line.CreateBound(center, center + XYZ(0, 0, 1))
                    try:
                        ElementTransformUtils.RotateElement(doc, hanger.Id, z_axis, rotation_angle)
                    except Exception as ex:
                        print("Error rotating hanger {}: {}".format(hanger.Id, str(ex)))

                    doc.Regenerate()

                    along_distance = min_along + current_spacing
                    target_position = reference_point + (direction * along_distance) + (perpendicular * trapeze_center_offset) + XYZ(0, 0, placement_z)
                    current_spacing += spacing

                    center = get_center_point(hanger.Id)
                    translation = target_position - center
                    try:
                        ElementTransformUtils.MoveElement(doc, hanger.Id, translation)
                    except Exception as ex:
                        print("Error moving hanger {}: {}".format(hanger.Id, str(ex)))

                    try:
                        offset_param = hanger.get_Parameter(BuiltInParameter.FABRICATION_OFFSET_PARAM)
                        if offset_param:
                            level_elevation = get_level_elevation_from_element(hanger)
                            offset_param.Set(placement_z - level_elevation)
                    except Exception as ex:
                        print("Error setting offset for hanger {}: {}".format(hanger.Id, str(ex)))

                    if attach_to_structure:
                        try:
                            hanger.GetRodInfo().AttachToStructure()
                        except Exception as ex:
                            print("Error attaching hanger {} to structure: {}".format(hanger.Id, str(ex)))

                if created_count == 0:
                    transaction.RollBack()
                    raise Exception("No hangers were created. Check fabrication service and button compatibility.")

                transaction.Commit()

            except Exception:
                transaction.RollBack()
                raise

    # --------------------------------------------------------------------------------
    # P-TRAP
    else:
        level_id = element.LevelId

        if selected_element_service_name in service_names:
            settings_lines[3] = selected_element_service_name

        dialog = HangerSpacingDialog(
            hanger_button_names,
            service_names,
            settings_lines,
            checkbox_boi_default,
            checkbox_attach_default,
            is_ptrap=True
        )

        if dialog.ShowDialog():
            selected_hanger_name = str(dialog.combobox_hanger.SelectedItem)
            attach_to_structure = dialog.checkbox_attach.IsChecked
            selected_service_name = str(dialog.combobox_service.SelectedItem)

            trap_width_text = dialog.textbox_trap_width.Text
            try:
                trap_width = float(trap_width_text)
                if trap_width <= 0:
                    raise ValueError("Width must be positive.")
            except ValueError:
                TaskDialog.Show("Invalid Input", "Trapeze Rod - Rod Width must be a positive number.")
                raise

            hanger_button = find_hanger_button(loaded_services, selected_service_name, selected_hanger_name)
            if not hanger_button:
                raise Exception("Hanger button '{}' not found in '{}'".format(selected_hanger_name, selected_service_name))

            write_settings_file(
                settings_path,
                selected_hanger_name,
                settings_lines[1],
                settings_lines[2],
                selected_service_name,
                attach_to_structure,
                settings_lines[5]
            )

            ptrap_bbox = element.get_BoundingBox(curview)
            if not ptrap_bbox:
                raise Exception("P-Trap bounding box not found.")

            insulation_thickness = element.InsulationThickness if hasattr(element, 'InsulationThickness') and element.HasInsulation else 0.0
            center_xy = (ptrap_bbox.Max + ptrap_bbox.Min) / 2
            bottom_z = ptrap_bbox.Min.Z - insulation_thickness if element.HasInsulation else ptrap_bbox.Min.Z
            target_position = XYZ(center_xy.X, center_xy.Y, bottom_z)

            connector_manager = element.ConnectorManager
            connector_2 = None
            connector_3 = None

            for connector in connector_manager.Connectors:
                if connector.Id == 1:
                    connector_2 = connector
                elif connector.Id == 2:
                    connector_3 = connector

            if not (connector_2 and connector_3):
                raise Exception("Required connectors not found.")

            connector_2_pos = connector_2.Origin
            connector_3_pos = connector_3.Origin
            direction = (connector_3_pos - connector_2_pos).Normalize()
            direction = XYZ(direction.X, direction.Y, 0).Normalize()
            rotation_angle = math.atan2(direction.Y, direction.X) + math.pi

            transaction = Transaction(doc, 'Place Trapeze Hanger on P-Trap')
            transaction.Start()

            try:
                hanger = FabricationPart.CreateHanger(doc, hanger_button, 0, level_id)
                if not hanger:
                    transaction.RollBack()
                    raise Exception("Hanger creation failed.")

                for dim in hanger.GetDimensions():
                    dim_name = dim.Name
                    try:
                        if dim_name == "Width":
                            hanger.SetDimensionValue(dim, (trap_width - 0.166666))
                        elif dim_name == "Bearer Extn":
                            hanger.SetDimensionValue(dim, 0.25)
                    except Exception as ex:
                        print("Error setting dimension '{}' for hanger {}: {}".format(dim_name, hanger.Id, str(ex)))

                center = get_center_point(hanger.Id)
                z_axis = Line.CreateBound(center, center + XYZ(0, 0, 1))
                try:
                    ElementTransformUtils.RotateElement(doc, hanger.Id, z_axis, rotation_angle)
                except Exception as ex:
                    transaction.RollBack()
                    raise Exception("Failed to rotate hanger: {}".format(str(ex)))

                center = get_center_point(hanger.Id)
                translation = target_position - center
                try:
                    ElementTransformUtils.MoveElement(doc, hanger.Id, translation)
                except Exception as ex:
                    transaction.RollBack()
                    raise Exception("Failed to move hanger: {}".format(str(ex)))

                try:
                    offset_param = hanger.get_Parameter(BuiltInParameter.FABRICATION_OFFSET_PARAM)
                    if offset_param:
                        level_elevation = get_level_elevation_from_element(hanger)
                        offset_param.Set(bottom_z - level_elevation)
                except Exception as ex:
                    print("Error setting offset for hanger {}: {}".format(hanger.Id, str(ex)))

                if attach_to_structure:
                    try:
                        hanger.GetRodInfo().AttachToStructure()
                    except Exception as ex:
                        print("Error attaching hanger {} to structure: {}".format(hanger.Id, str(ex)))

                transaction.Commit()

            except Exception:
                if transaction.HasStarted():
                    transaction.RollBack()
                raise

except Exception as ex:
    TaskDialog.Show("Error", "Script error: {}".format(str(ex)))