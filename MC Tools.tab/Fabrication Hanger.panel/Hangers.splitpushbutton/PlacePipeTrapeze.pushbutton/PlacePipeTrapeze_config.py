# -*- coding: utf-8 -*-
import Autodesk
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.DB import Transaction, FabricationConfiguration, BuiltInParameter, FabricationPart, \
                                XYZ, ElementTransformUtils, Line, BuiltInCategory, Plane, SketchPlane
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType, ObjectSnapTypes
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
# DEFINE SOME VARIABLES EASY USE
# ------------------------------------------------------------------------------------
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

# ------------------------------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------------------------------
def get_id_value(id_obj):
    try:
        return id_obj.Value
    except:
        return id_obj.IntegerValue

def ensure_active_work_plane(seed_point=None):
    try:
        if curview.SketchPlane:
            return
    except:
        pass

    t = Transaction(doc, "Set Temporary Work Plane")
    t.Start()
    try:
        origin = seed_point
        if origin is None:
            try:
                origin = curview.Origin
            except:
                origin = XYZ(0, 0, 0)

        plane = Plane.CreateByNormalAndOrigin(curview.ViewDirection, origin)
        sp = SketchPlane.Create(doc, plane)
        curview.SketchPlane = sp
        t.Commit()
    except:
        t.RollBack()
        raise Exception("A work plane is required for point picking. Set a work plane in the current view and rerun.")

def pick_width_points(seed_point=None):
    ensure_active_work_plane(seed_point)

    snap_types = (
        ObjectSnapTypes.Endpoints |
        ObjectSnapTypes.Midpoints |
        ObjectSnapTypes.Intersections |
        ObjectSnapTypes.Nearest
    )

    p1 = uidoc.Selection.PickPoint(
        snap_types,
        "Pick FIRST trapeze width point - snap/track to structural framing line"
    )
    p2 = uidoc.Selection.PickPoint(
        snap_types,
        "Pick SECOND trapeze width point - snap/track to structural framing line"
    )
    return p1, p2

class FabricationPipeSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        try:
            return element.Category and get_id_value(element.Category.Id) == int(BuiltInCategory.OST_FabricationPipework)
        except:
            return False

    def AllowReference(self, reference, position):
        return False

# ------------------------------------------------------------------------------------
# SELECTING ELEMENTS
# ------------------------------------------------------------------------------------
try:
    selected_element = uidoc.Selection.PickObject(ObjectType.Element, 'Select OUTSIDE Pipe')
    element = doc.GetElement(selected_element.ElementId)
    pick_point = selected_element.GlobalPoint

    parameters = element.LookupParameter('Fabrication Service')
    if parameters and parameters.HasValue:
        service_name = parameters.AsValueString()
    else:
        raise Exception("Fabrication Service parameter missing.")

    servicenamelist = []
    Config = FabricationConfiguration.GetFabricationConfiguration(doc)
    LoadedServices = Config.GetAllLoadedServices()
    for Item1 in LoadedServices:
        try:
            servicenamelist.append(Item1.Name)
        except:
            servicenamelist.append([])

    try:
        Servicenum = servicenamelist.index(service_name)
    except ValueError:
        raise Exception("Selected service not found.")

    buttonnames = []
    unique_hangers = set()
    for service_idx, service in enumerate(LoadedServices):
        palette_count = service.PaletteCount if RevitINT > 2022 else service.GroupCount
        for palette_idx in range(palette_count):
            buttoncount = service.GetButtonCount(palette_idx)
            for btn_idx in range(buttoncount):
                bt = service.GetButton(palette_idx, btn_idx)
                if bt.IsAHanger and bt.Name not in unique_hangers:
                    unique_hangers.add(bt.Name)
                    buttonnames.append(bt.Name)

    folder_name = "c:\\Temp"
    filepath = os.path.join(folder_name, 'Ribbon_PlaceTrapeze.txt')
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)

    defaults = [
        '1.625 Single Strut Trapeze\n',
        '1.0\n',
        '8.0\n',
        'PLUMBING: DOMESTIC COLD WATER\n',
        'True\n',
        'True'
    ]

    if not os.path.exists(filepath):
        with open(filepath, 'w') as the_file:
            the_file.writelines(defaults)

    with open(filepath, 'r') as file:
        lines = [line.rstrip() for line in file.readlines()]

    if len(lines) < 6:
        with open(filepath, 'w') as the_file:
            the_file.writelines(defaults)
        lines = [x.rstrip() for x in defaults]

    checkboxdef = lines[4] != 'False'
    checkboxdefBOI = lines[5] != 'False'

    # --------------------------------------------------------------------------------
    # DIALOG
    # --------------------------------------------------------------------------------
    class HangerSpacingDialog(Window):
        def __init__(self, buttonnames, lines, checkboxdefBOI, checkboxdef, is_ptrap=False):
            super(HangerSpacingDialog, self).__init__()
            self.Title = "Hanger and Spacing" if not is_ptrap else "Hanger for P-Trap"
            self.Width = 336
            self.Height = 350 if not is_ptrap else 275
            self.WindowStartupLocation = WindowStartupLocation.CenterScreen
            self.ResizeMode = ResizeMode.NoResize
            self.is_ptrap = is_ptrap

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
            self.combobox_hanger.ItemsSource = Array[object](buttonnames)
            if lines[0] in buttonnames:
                self.combobox_hanger.SelectedItem = lines[0]
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
                self.textbox_end_dist.Text = str(round(float(lines[1]) * 12.0, 4))
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
            self.combobox_service.ItemsSource = Array[object](servicenamelist)
            if lines[3] in servicenamelist:
                self.combobox_service.SelectedItem = lines[3]
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

    # --------------------------------------------------------------------------------
    # REGULAR TRAPEZE CASE
    # --------------------------------------------------------------------------------
    if element.ItemCustomId != 916:
        rack_refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            FabricationPipeSelectionFilter(),
            'Window select rack fabrication pipes'
        )

        selected_elements = []
        seen_ids = set()

        selected_elements.append(element)
        seen_ids.add(get_id_value(element.Id))

        for r in rack_refs:
            p = doc.GetElement(r.ElementId)
            if p and get_id_value(p.Id) not in seen_ids:
                selected_elements.append(p)
                seen_ids.add(get_id_value(p.Id))

        if len(selected_elements) == 0:
            raise Exception("No rack pipes selected.")

        # Pick width points instead of structural members
        width_pick_1, width_pick_2 = pick_width_points(pick_point)

        level_id = element.LevelId

        def get_parameter_value(element, parameterName):
            param = element.LookupParameter(parameterName)
            if param and param.HasValue:
                return param.AsDouble()
            else:
                return 0.0

        if RevitINT > 2022:
            PRTElevation = get_parameter_value(element, 'Lower End Bottom Elevation')
        else:
            PRTElevation = get_parameter_value(element, 'Bottom')

        form = HangerSpacingDialog(buttonnames, lines, checkboxdefBOI, checkboxdef, is_ptrap=False)
        if form.ShowDialog():
            Selectedbutton = str(form.combobox_hanger.SelectedItem)
            distancefromend = form.textbox_end_dist.Text
            Spacing = form.textbox_spacing.Text
            BOITrap = form.checkbox_boi.IsChecked
            AtoS = form.checkbox_attach.IsChecked
            SelectedServiceName = str(form.combobox_service.SelectedItem)

            try:
                distancefromend = float(distancefromend) / 12.0
                Spacing = float(Spacing)
                if Spacing <= 0:
                    raise ValueError
                if distancefromend < 0:
                    raise ValueError
            except ValueError:
                raise Exception("Distance from End must be >= 0 and Spacing must be > 0.")

            try:
                Servicenum = servicenamelist.index(SelectedServiceName)
            except ValueError:
                raise Exception("Selected service not found.")

            button_found = False
            fab_btn = None
            for servicenum, service in enumerate(LoadedServices):
                if service.Name == SelectedServiceName:
                    palette_count = service.PaletteCount if RevitINT > 2022 else service.GroupCount
                    for palette_idx in range(palette_count):
                        button_count = service.GetButtonCount(palette_idx)
                        for btn_idx in range(button_count):
                            bt = service.GetButton(palette_idx, btn_idx)
                            if bt.Name == Selectedbutton:
                                fab_btn = bt
                                button_found = True
                                break
                        if button_found:
                            break
                    if button_found:
                        break

            if not button_found:
                raise Exception("Hanger button not found.")

            with open(filepath, 'w') as the_file:
                the_file.writelines([
                    str(Selectedbutton) + '\n',
                    str(distancefromend) + '\n',
                    str(Spacing) + '\n',
                    SelectedServiceName + '\n',
                    str(AtoS) + '\n',
                    str(BOITrap) + '\n'
                ])

            def GetCenterPoint(ele_id):
                bBox = doc.GetElement(ele_id).get_BoundingBox(None)
                if bBox:
                    return (bBox.Max + bBox.Min) / 2
                return XYZ(0, 0, 0)

            def myround(x, multiple):
                return multiple * math.ceil(x / multiple)

            def get_diameter(pipe):
                param = pipe.LookupParameter("Outside Diameter")
                if param and param.HasValue:
                    return param.AsDouble()
                return 0.0

            def get_reference_level(hanger):
                return doc.GetElement(hanger.LevelId)

            def get_level_elevation(level):
                if level:
                    try:
                        return level.ProjectElevation
                    except:
                        try:
                            return level.Elevation
                        except:
                            return 0.0
                return 0.0

            curve = element.Location.Curve
            if not curve or not curve.IsBound:
                raise Exception("Pipe must have a valid location curve.")

            dir_vec = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize()
            dir_vec = XYZ(dir_vec.X, dir_vec.Y, 0).Normalize()
            perp_vec = XYZ(-dir_vec.Y, dir_vec.X, 0).Normalize()

            endpoints = []
            for pipe in selected_elements:
                c = pipe.Location.Curve
                if c and c.IsBound:
                    endpoints.append(c.GetEndPoint(0))
                    endpoints.append(c.GetEndPoint(1))

            projs_along = [p.DotProduct(dir_vec) for p in endpoints]
            min_along_global = min(projs_along)
            max_along_global = max(projs_along)
            length_along = max_along_global - min_along_global

            combined_min_z = float('inf')
            for pipe in selected_elements:
                pipe_bb = pipe.get_BoundingBox(None)
                if pipe_bb:
                    bottom_z = pipe_bb.Min.Z
                    thick = pipe.InsulationThickness if hasattr(pipe, 'InsulationThickness') and pipe.HasInsulation else 0.0
                    if BOITrap:
                        bottom_z -= thick
                    combined_min_z = min(combined_min_z, bottom_z)

            center_z = combined_min_z

            qtyofhgrs = int(math.ceil(length_along / Spacing))
            if qtyofhgrs < 1:
                qtyofhgrs = 1

            hangers = []
            t = Transaction(doc, 'Place Trapeze Hanger')
            t.Start()
            for hgr in range(qtyofhgrs):
                try:
                    hanger = FabricationPart.CreateHanger(doc, fab_btn, 0, level_id)
                    if hanger:
                        hangers.append(hanger)
                except:
                    pass
            t.Commit()

            if not hangers:
                raise Exception("Hanger creation failed.")

            t = Transaction(doc, 'Modify Trapeze Hanger')
            t.Start()
            IncrementSpacing = distancefromend

            first_pipe = element
            curve = first_pipe.Location.Curve
            if not curve or not curve.IsBound:
                raise Exception("First pipe must have a valid location curve.")

            first_endpoints = [curve.GetEndPoint(0), curve.GetEndPoint(1)]
            end0 = first_endpoints[0]
            end1 = first_endpoints[1]
            dist0 = end0.DistanceTo(pick_point)
            dist1 = end1.DistanceTo(pick_point)

            if dist0 <= dist1:
                ref_point = end0
                other_end = end1
            else:
                ref_point = end1
                other_end = end0

            dir_vec_for_placement = (other_end - ref_point).Normalize()
            dir_vec = XYZ(dir_vec_for_placement.X, dir_vec_for_placement.Y, 0).Normalize()
            perp_vec = XYZ(-dir_vec.Y, dir_vec.X, 0).Normalize()

            # Width and center come from the 2 picked points
            p1_perp = (width_pick_1 - ref_point).DotProduct(perp_vec)
            p2_perp = (width_pick_2 - ref_point).DotProduct(perp_vec)

            width = abs(p2_perp - p1_perp)
            center_perp = (p1_perp + p2_perp) / 2.0

            if width <= 0:
                t.RollBack()
                raise Exception("Picked width points do not create a valid trapeze width.")

            projs_along = [(p - ref_point).DotProduct(dir_vec) for p in first_endpoints]
            min_along = min(projs_along)

            for idx, hanger in enumerate(hangers):
                newwidth = myround(width * 12, 2) / 12
                for dim in hanger.GetDimensions():
                    dim_name = dim.Name
                    try:
                        if dim_name in ("Width", "Duct Width"):
                            hanger.SetDimensionValue(dim, newwidth)
                        if dim_name == "Bearer Extn":
                            hanger.SetDimensionValue(dim, 0.16666)
                    except:
                        pass

                center = GetCenterPoint(hanger.Id)
                z_axis = Line.CreateBound(center, center + XYZ(0, 0, 1))
                angle_rad = math.atan2(dir_vec.Y, dir_vec.X)
                try:
                    ElementTransformUtils.RotateElement(doc, hanger.Id, z_axis, angle_rad)
                except:
                    pass

                along = min_along + IncrementSpacing
                pos = ref_point + dir_vec * along + perp_vec * center_perp + XYZ(0, 0, center_z)
                IncrementSpacing += Spacing

                center = GetCenterPoint(hanger.Id)
                translation = pos - center
                try:
                    ElementTransformUtils.MoveElement(doc, hanger.Id, translation)
                except:
                    pass

                reference_level = get_reference_level(hanger)
                elevation = get_level_elevation(reference_level)
                try:
                    offset_param = hanger.get_Parameter(BuiltInParameter.FABRICATION_OFFSET_PARAM)
                    offset_value = center_z - elevation
                    offset_param.Set(offset_value)
                except:
                    pass

                if AtoS:
                    try:
                        hanger.GetRodInfo().AttachToStructure()
                    except Exception as e:
                        TaskDialog.Show("Error", "Error attaching hanger {} to structure: {}".format(hanger.Id, str(e)))

            t.Commit()

    # --------------------------------------------------------------------------------
    # P-TRAP / CID 916 CASE
    # --------------------------------------------------------------------------------
    else:
        level_id = element.LevelId
        if service_name in servicenamelist:
            lines[3] = service_name

        form = HangerSpacingDialog(buttonnames, lines, checkboxdefBOI, checkboxdef, is_ptrap=True)
        if form.ShowDialog():
            Selectedbutton = str(form.combobox_hanger.SelectedItem)
            AtoS = form.checkbox_attach.IsChecked
            SelectedServiceName = str(form.combobox_service.SelectedItem)

            trap_width_text = form.textbox_trap_width.Text
            try:
                trap_width = float(trap_width_text)
                if trap_width <= 0:
                    raise ValueError("Width must be positive.")
            except ValueError:
                TaskDialog.Show("Invalid Input", "Trapeze Rod - Rod Width must be a positive number.")
                raise

            try:
                Servicenum = servicenamelist.index(SelectedServiceName)
            except ValueError:
                raise Exception("Selected service not found.")

            button_found = False
            fab_btn = None
            for servicenum, service in enumerate(LoadedServices):
                if service.Name == SelectedServiceName:
                    palette_count = service.PaletteCount if RevitINT > 2022 else service.GroupCount
                    for palette_idx in range(palette_count):
                        button_count = service.GetButtonCount(palette_idx)
                        for btn_idx in range(button_count):
                            bt = service.GetButton(palette_idx, btn_idx)
                            if bt.Name == Selectedbutton:
                                fab_btn = bt
                                button_found = True
                                break
                        if button_found:
                            break
                    if button_found:
                        break

            if not button_found:
                raise Exception("Hanger button not found.")

            with open(filepath, 'w') as the_file:
                the_file.writelines([
                    str(Selectedbutton) + '\n',
                    lines[1] + '\n',
                    lines[2] + '\n',
                    SelectedServiceName + '\n',
                    str(AtoS) + '\n',
                    lines[5] + '\n'
                ])

            def GetCenterPoint(ele_id):
                bBox = doc.GetElement(ele_id).get_BoundingBox(None)
                if bBox:
                    return (bBox.Max + bBox.Min) / 2
                return XYZ(0, 0, 0)

            def get_reference_level(hanger):
                return doc.GetElement(hanger.LevelId)

            def get_level_elevation(level):
                if level:
                    try:
                        return level.Elevation
                    except:
                        return 0.0
                return 0.0

            ptrap_bb = element.get_BoundingBox(curview)
            if not ptrap_bb:
                raise Exception("P-Trap bounding box not found.")

            thick = element.InsulationThickness if hasattr(element, 'InsulationThickness') and element.HasInsulation else 0.0
            center_xy = (ptrap_bb.Max + ptrap_bb.Min) / 2
            bottom_z = ptrap_bb.Min.Z - thick if element.HasInsulation else ptrap_bb.Min.Z
            target_pos = XYZ(center_xy.X, center_xy.Y, bottom_z)

            connector_manager = element.ConnectorManager
            c2 = None
            c3 = None
            for connector in connector_manager.Connectors:
                if connector.Id == 1:
                    c2 = connector
                elif connector.Id == 2:
                    c3 = connector

            if not (c2 and c3):
                raise Exception("Required connectors not found.")

            c2_pos = c2.Origin
            c3_pos = c3.Origin
            dir_vec = (c3_pos - c2_pos).Normalize()
            dir_vec = XYZ(dir_vec.X, dir_vec.Y, 0).Normalize()
            angle_rad = math.atan2(dir_vec.Y, dir_vec.X) + math.pi

            t = Transaction(doc, 'Place Trapeze Hanger on P-Trap')
            t.Start()
            try:
                hanger = FabricationPart.CreateHanger(doc, fab_btn, 0, level_id)
                if not hanger:
                    t.RollBack()
                    raise Exception("Hanger creation failed.")
            except Exception as e:
                t.RollBack()
                raise Exception("Hanger creation failed: {}".format(str(e)))

            for dim in hanger.GetDimensions():
                dim_name = dim.Name
                try:
                    if dim_name == "Width":
                        hanger.SetDimensionValue(dim, (trap_width - 0.166666))
                    if dim_name == "Bearer Extn":
                        hanger.SetDimensionValue(dim, 0.25)
                except:
                    pass

            center = GetCenterPoint(hanger.Id)
            z_axis = Line.CreateBound(center, center + XYZ(0, 0, 1))
            try:
                ElementTransformUtils.RotateElement(doc, hanger.Id, z_axis, angle_rad)
            except Exception as e:
                t.RollBack()
                raise Exception("Failed to rotate hanger: {}".format(str(e)))

            center = GetCenterPoint(hanger.Id)
            translation = target_pos - center
            try:
                ElementTransformUtils.MoveElement(doc, hanger.Id, translation)
            except Exception as e:
                t.RollBack()
                raise Exception("Failed to move hanger: {}".format(str(e)))

            reference_level = get_reference_level(hanger)
            elevation = get_level_elevation(reference_level)
            try:
                offset_param = hanger.get_Parameter(BuiltInParameter.FABRICATION_OFFSET_PARAM)
                offset_value = bottom_z - elevation
                offset_param.Set(offset_value)
            except:
                pass

            if AtoS:
                try:
                    hanger.GetRodInfo().AttachToStructure()
                except Exception as e:
                    print("Error attaching hanger {} to structure: {}".format(hanger.Id, str(e)))

            t.Commit()

except Exception as e:
    TaskDialog.Show("Error", "Script error: {}".format(str(e)))