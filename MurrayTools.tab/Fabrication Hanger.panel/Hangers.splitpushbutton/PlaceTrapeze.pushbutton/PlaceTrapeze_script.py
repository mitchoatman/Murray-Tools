import Autodesk
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.DB import Transaction, FabricationConfiguration, BuiltInParameter, FabricationPart, FabricationServiceButton, \
                                FabricationService, XYZ, ElementTransformUtils, BoundingBoxXYZ, Transform, Line
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
#------------------------------------------------------------------------------------DEFINE SOME VARIABLES EASY USE
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)
#------------------------------------------------------------------------------------SELECTING ELEMENTS
try:
    selected_element = uidoc.Selection.PickObject(ObjectType.Element, 'Select OUTSIDE Pipe')
    element = doc.GetElement(selected_element.ElementId)
    pick_point = selected_element.GlobalPoint  # Capture cursor location for nearest end determination
 
    # Gets servicename of selection (used in both cases)
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
    # Gets matching index of selected element service
    try:
        Servicenum = servicenamelist.index(service_name)
    except ValueError:
        raise Exception("Selected service not found.")
    # Find all hanger buttons
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
    if not os.path.exists(filepath):
        with open(filepath, 'w') as the_file:
            lines = ['1.625 Single Strut Trapeze\n', '1.0\n', '8.0\n', 'PLUMBING: DOMESTIC COLD WATER\n', 'True\n', 'True']
            the_file.writelines(lines)
    with open(filepath, 'r') as file:
        lines = [line.rstrip() for line in file.readlines()]
    if len(lines) < 6:
        with open(filepath, 'w') as the_file:
            lines = ['1.625 Single Strut Trapeze\n', '1.0\n', '8.0\n', 'PLUMBING: DOMESTIC COLD WATER\n', 'True\n', 'True']
            the_file.writelines(lines)
    with open(filepath, 'r') as file:
        lines = [line.rstrip() for line in file.readlines()]
    checkboxdef = lines[4] != 'False'
    checkboxdefBOI = lines[5] != 'False'
    # Define HangerSpacingDialog (already exists, reused for both cases)
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
            # Choose Hanger Label
            label_hanger = Label()
            label_hanger.Content = "Choose Hanger:"
            label_hanger.FontSize = 12
            label_hanger.FontFamily = FontFamily("Arial")
            label_hanger.Margin = Thickness(0, 0, 0, 0)
            stack.Children.Add(label_hanger)
            # Hanger ComboBox
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
                # Distance from End Label
                label_end_dist = Label()
                label_end_dist.Content = "Distance from End (In):"
                label_end_dist.FontSize = 12
                label_end_dist.FontFamily = FontFamily("Arial")
                label_end_dist.Margin = Thickness(0, 0, 0, 0)
                stack.Children.Add(label_end_dist)
                # Distance from End TextBox
                self.textbox_end_dist = TextBox()
                self.textbox_end_dist.Width = 200
                self.textbox_end_dist.Height = 20
                self.textbox_end_dist.FontSize = 12
                self.textbox_end_dist.FontFamily = FontFamily("Arial")
                self.textbox_end_dist.Text = str(round(float(lines[1]) * 12.0, 4))
                self.textbox_end_dist.Margin = Thickness(0, 0, 0, 10)
                self.textbox_end_dist.HorizontalAlignment = HorizontalAlignment.Left
                stack.Children.Add(self.textbox_end_dist)
                # Hanger Spacing Label
                label_spacing = Label()
                label_spacing.Content = "Hanger Spacing (Ft):"
                label_spacing.FontSize = 12
                label_spacing.FontFamily = FontFamily("Arial")
                label_spacing.Margin = Thickness(0, 0, 0, 0)
                stack.Children.Add(label_spacing)
                # Hanger Spacing TextBox
                self.textbox_spacing = TextBox()
                self.textbox_spacing.Width = 200
                self.textbox_spacing.Height = 20
                self.textbox_spacing.FontSize = 12
                self.textbox_spacing.FontFamily = FontFamily("Arial")
                self.textbox_spacing.Text = lines[2]
                self.textbox_spacing.Margin = Thickness(0, 0, 0, 10)
                self.textbox_spacing.HorizontalAlignment = HorizontalAlignment.Left
                stack.Children.Add(self.textbox_spacing)
                # Align Trapeze to Bottom of Insulation CheckBox (only for regular trapeze)
                self.checkbox_boi = CheckBox()
                self.checkbox_boi.Content = "Align Trapeze to Bottom of Insulation"
                self.checkbox_boi.FontSize = 12
                self.checkbox_boi.FontFamily = FontFamily("Arial")
                self.checkbox_boi.IsChecked = checkboxdefBOI
                self.checkbox_boi.Margin = Thickness(0, 0, 0, 5)
                stack.Children.Add(self.checkbox_boi)
            # Attach to Structure CheckBox
            self.checkbox_attach = CheckBox()
            self.checkbox_attach.Content = "Attach to Structure"
            self.checkbox_attach.FontSize = 12
            self.checkbox_attach.FontFamily = FontFamily("Arial")
            self.checkbox_attach.IsChecked = checkboxdef
            self.checkbox_attach.Margin = Thickness(0, 0, 0, 5 if is_ptrap else 10)
            stack.Children.Add(self.checkbox_attach)
            # Trapeze width input for P-Trap only
            if is_ptrap:
                label_trap_width = Label()
                label_trap_width.Content = "Trapeze Rod - Rod Width (Ft):"
                label_trap_width.FontSize = 12
                label_trap_width.FontFamily = FontFamily("Arial")
                label_trap_width.Margin = Thickness(0, 0, 0, 0)
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
            # Choose Service Label
            label_service = Label()
            label_service.Content = "Choose Service to Draw Hanger on:"
            label_service.FontSize = 12
            label_service.FontFamily = FontFamily("Arial")
            label_service.Margin = Thickness(0, 0, 0, 0)
            stack.Children.Add(label_service)
            # Service ComboBox
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
            # OK Button
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
    if element.ItemCustomId != 916:
        selected_element1 = uidoc.Selection.PickObject(ObjectType.Element, 'Select OPPOSITE OUTSIDE Pipe')
        element1 = doc.GetElement(selected_element1.ElementId)
        selected_elements = [element, element1]
        level_id = element.LevelId
        # FUNCTION TO GET PARAMETER VALUE
        def get_parameter_value(element, parameterName):
            param = element.LookupParameter(parameterName)
            if param and param.HasValue:
                return param.AsDouble()
            else:
                return 0.0
        # Gets bottom elevation of selected pipe (no longer used after fixes)
        if RevitINT > 2022:
            PRTElevation = get_parameter_value(element, 'Lower End Bottom Elevation')
        else:
            PRTElevation = get_parameter_value(element, 'Bottom')
        # Instantiate and show the dialog
        form = HangerSpacingDialog(buttonnames, lines, checkboxdefBOI, checkboxdef, is_ptrap=False)
        if form.ShowDialog():
            Selectedbutton = str(form.combobox_hanger.SelectedItem)
            distancefromend = form.textbox_end_dist.Text
            Spacing = form.textbox_spacing.Text
            BOITrap = form.checkbox_boi.IsChecked
            AtoS = form.checkbox_attach.IsChecked
            SelectedServiceName = str(form.combobox_service.SelectedItem)
         
            # Validate numeric inputs
            try:
                distancefromend = float(distancefromend) / 12.0
                Spacing = float(Spacing)
            except ValueError:
                print("Invalid input: Distance from End or Spacing must be numeric.")
                raise Exception("Invalid numeric input.")
         
            # Gets matching index of selected service
            try:
                Servicenum = servicenamelist.index(SelectedServiceName)
            except ValueError:
                print("Selected service '{}' not found.".format(SelectedServiceName))
                raise Exception("Selected service not found.")
         
            # Find the selected button
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
                print("'{}' not found in '{}'".format(Selectedbutton, SelectedServiceName))
                raise Exception("Hanger button not found.")
         
            # Write values to text file
            with open(filepath, 'w') as the_file:
                the_file.writelines([
                    str(Selectedbutton) + '\n',
                    str(distancefromend) + '\n',
                    str(Spacing) + '\n',
                    SelectedServiceName + '\n',
                    str(AtoS) + '\n',
                    str(BOITrap) + '\n'
                ])
         
            # Helper functions
            def GetCenterPoint(ele_id):
                bBox = doc.GetElement(ele_id).get_BoundingBox(None)
                if bBox:
                    center = (bBox.Max + bBox.Min) / 2
                    return center
                else:
                    print("No bounding box found for element {}".format(ele_id))
                    return XYZ(0, 0, 0)
         
            def myround(x, multiple):
                return multiple * math.ceil(x / multiple)
         
            def get_diameter(pipe):
                param = pipe.LookupParameter("Outside Diameter")
                if param and param.HasValue:
                    return param.AsDouble()
                print("Outside Diameter parameter not found for pipe {}. Assuming 0.".format(pipe.Id))
                return 0.0
         
            def get_reference_level(hanger):
                level_id = hanger.LevelId
                level = doc.GetElement(level_id)
                return level
         
            def get_level_elevation(level):
                if level:
                    return level.Elevation
                return 0.0
         
            # Determine pipe direction from the first pipe's location curve
            curve = element.Location.Curve
            if not curve or not curve.IsBound:
                print("Invalid location curve for element {}".format(element.Id))
                raise Exception("Pipe must have a valid location curve.")
         
            # Get direction vector in XY plane (ignore Z for rotation on Z-axis)
            dir_vec = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize()
            dir_vec = XYZ(dir_vec.X, dir_vec.Y, 0).Normalize() # Ensure Z is 0
            perp_vec = XYZ(-dir_vec.Y, dir_vec.X, 0).Normalize() # Perpendicular vector in XY plane
         
            # Compute midpoints of both pipes
            midpoints = []
            for pipe in selected_elements:
                c = pipe.Location.Curve
                if c and c.IsBound:
                    mid = (c.GetEndPoint(0) + c.GetEndPoint(1)) / 2
                    midpoints.append(mid)
                else:
                    print("Invalid location curve for pipe {}".format(pipe.Id))
                    raise Exception("Pipe must have a valid location curve.")
         
            # Compute center point between the two pipes in the perpendicular direction
            perp_projs = [mid.DotProduct(perp_vec) for mid in midpoints]
            min_perp = min(perp_projs)
            max_perp = max(perp_projs)
            center_perp = (min_perp + max_perp) / 2
         
            # Adjust for pipe diameter and insulation
            widths = []
            for pipe in selected_elements:
                thick = pipe.InsulationThickness if pipe.HasInsulation else 0
                half_size = get_diameter(pipe) / 2 + thick
                widths.append(half_size)
         
            # Compute effective width (distance between outermost edges)
            width = abs(max_perp - min_perp) + widths[0] + widths[1]
         
            # Compute projections along direction to find length
            endpoints = []
            for pipe in selected_elements:
                c = pipe.Location.Curve
                if c and c.IsBound:
                    endpoints.append(c.GetEndPoint(0))
                    endpoints.append(c.GetEndPoint(1))
         
            projs_along = [p.DotProduct(dir_vec) for p in endpoints]
            min_along = min(projs_along)
            max_along = max(projs_along)
            length_along = max_along - min_along
         
            # Compute lowest bottom Z across both pipes using bounding box
            combined_min_z = float('inf')
            for pipe in selected_elements:
                pipe_bb = pipe.get_BoundingBox(curview)
                if pipe_bb:
                    bottom_z = pipe_bb.Min.Z
                    thick = pipe.InsulationThickness if hasattr(pipe, 'InsulationThickness') and pipe.HasInsulation else 0.0
                    if BOITrap:
                        bottom_z -= thick # Adjust only when checkbox is checked
                    combined_min_z = min(combined_min_z, bottom_z)
            center_z = combined_min_z # Always use the lowest adjusted bottom
         
            # Calculate number of hangers
            qtyofhgrs = int(math.ceil(length_along / Spacing))
         
            # Place hangers at default location (0,0,0)
            hangers = []
            t = Transaction(doc, 'Place Trapeze Hanger')
            t.Start()
            for hgr in range(qtyofhgrs):
                try:
                    hanger = FabricationPart.CreateHanger(doc, fab_btn, 0, level_id)
                    if hanger:
                        hangers.append(hanger)
                    else:
                        print("Failed to create hanger {}".format(hgr + 1))
                except Exception as e:
                    print("Error creating hanger {}: {}".format(hgr + 1, str(e)))
            t.Commit()
         
            if not hangers:
                print("No hangers were created. Check fabrication service and button compatibility.")
                raise Exception("Hanger creation failed.")
         
            # Move and modify hangers
            t = Transaction(doc, 'Modify Trapeze Hanger')
            t.Start()
            IncrementSpacing = distancefromend
         
            # Find the closest endpoint of the first selected pipe to the pick point (cursor location)
            first_pipe = selected_elements[0]
            curve = first_pipe.Location.Curve
            if not curve or not curve.IsBound:
                print("Invalid location curve for first pipe {}".format(first_pipe.Id))
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
            dir_vec = XYZ(dir_vec_for_placement.X, dir_vec_for_placement.Y, 0).Normalize() # Ensure Z=0, directed away from ref_point
            perp_vec = XYZ(-dir_vec.Y, dir_vec.X, 0).Normalize() # Perpendicular in XY plane
         
            # Compute midpoints, diameters, and insulation for both pipes
            endpoints = []
            midpoints = []
            thicknesses = []
            diameters = []
            for pipe in selected_elements:
                c = pipe.Location.Curve
                if c and c.IsBound:
                    endpoints.extend([c.GetEndPoint(0), c.GetEndPoint(1)])
                    mid = (c.GetEndPoint(0) + c.GetEndPoint(1)) / 2
                    midpoints.append(mid)
                    thick = pipe.InsulationThickness if hasattr(pipe, 'InsulationThickness') and pipe.HasInsulation else 0.0
                    thicknesses.append(thick)
                    diam = get_diameter(pipe)
                    diameters.append(diam)
                else:
                    print("Invalid location curve for pipe {}".format(pipe.Id))
                    raise Exception("Pipe must have a valid location curve.")
         
            # Recalculate center_perp using projections onto perp_vec
            perp_projs = []
            for i, mid in enumerate(midpoints):
                proj = (mid - ref_point).DotProduct(perp_vec)
                half_size = diameters[i] / 2 + thicknesses[i]
                perp_projs.append(proj - half_size) # Min edge
                perp_projs.append(proj + half_size) # Max edge
            center_perp = (min(perp_projs) + max(perp_projs)) / 2
         
            # Recalculate projections along the pipe direction for first pipe
            projs_along = [(p - ref_point).DotProduct(dir_vec) for p in first_endpoints]
            min_along = min(projs_along)
         
            for idx, hanger in enumerate(hangers):
                # Set dimensions first
                newwidth = myround(width * 12, 2) / 12
                for dim in hanger.GetDimensions():
                    dim_name = dim.Name
                    try:
                        if dim_name == "Width":
                            hanger.SetDimensionValue(dim, newwidth)
                        if dim_name == "Bearer Extn":
                            hanger.SetDimensionValue(dim, 0.25)
                    except Exception as e:
                        print("Error setting dimension '{}' for hanger {}: {}".format(dim_name, hanger.Id, str(e)))
             
                # Rotate hanger to align with pipe direction
                center = GetCenterPoint(hanger.Id)
                z_axis = Line.CreateBound(center, center + XYZ(0, 0, 1))
                angle_rad = math.atan2(dir_vec.Y, dir_vec.X)
                try:
                    ElementTransformUtils.RotateElement(doc, hanger.Id, z_axis, angle_rad)
                except Exception as e:
                    print("Error rotating hanger {}: {}".format(hanger.Id, str(e)))
             
                # Compute target position starting from nearest endpoint
                along = min_along + IncrementSpacing
                pos = ref_point + dir_vec * along + perp_vec * center_perp + XYZ(0, 0, center_z)
                IncrementSpacing += Spacing
             
                # Move hanger
                center = GetCenterPoint(hanger.Id)
                translation = pos - center
                try:
                    ElementTransformUtils.MoveElement(doc, hanger.Id, translation)
                except Exception as e:
                    print("Error moving hanger {}: {}".format(hanger.Id, str(e)))
             
                # Set offset
                reference_level = get_reference_level(hanger)
                elevation = get_level_elevation(reference_level)
                try:
                    offset_param = hanger.get_Parameter(BuiltInParameter.FABRICATION_OFFSET_PARAM)
                    offset_value = center_z - elevation # Always relative, using bounding-box-derived bottom
                    offset_param.Set(offset_value)
                except Exception as e:
                    print("Error setting offset for hanger {}: {}".format(hanger.Id, str(e)))
             
                if AtoS:
                    try:
                        hanger.GetRodInfo().AttachToStructure()
                    except Exception as e:
                        TaskDialog.Show("Error", "Error attaching hanger {} to structure: {}".format(hanger.Id, str(e)))
         
            t.Commit()
     
        else:
            pass
 
    else:
        # Handle CID 916 (P-Trap) case
        level_id = element.LevelId
        if service_name in servicenamelist:
            lines[3] = service_name
        # Instantiate and show the dialog (reusing HangerSpacingDialog with is_ptrap=True)
        form = HangerSpacingDialog(buttonnames, lines, checkboxdefBOI, checkboxdef, is_ptrap=True)
        if form.ShowDialog():
            Selectedbutton = str(form.combobox_hanger.SelectedItem)
            AtoS = form.checkbox_attach.IsChecked
            SelectedServiceName = str(form.combobox_service.SelectedItem)
            # Get and validate user-defined trapeze width
            trap_width_text = form.textbox_trap_width.Text
            try:
                trap_width = float(trap_width_text)
                if trap_width <= 0:
                    raise ValueError("Width must be positive.")
            except ValueError:
                TaskDialog.Show("Invalid Input", "Trapeze Rod - Rod Width must be a positive number.")
                raise
            # Gets matching index of selected service
            try:
                Servicenum = servicenamelist.index(SelectedServiceName)
            except ValueError:
                print("Selected service '{}' not found.".format(SelectedServiceName))
                raise Exception("Selected service not found.")
         
            # Find the selected button
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
                print("'{}' not found in '{}'".format(Selectedbutton, SelectedServiceName))
                raise Exception("Hanger button not found.")
         
            # Write values to text file (preserve distance/spacing/BOI from previous values)
            with open(filepath, 'w') as the_file:
                the_file.writelines([
                    str(Selectedbutton) + '\n',
                    lines[1] + '\n',
                    lines[2] + '\n',
                    SelectedServiceName + '\n',
                    str(AtoS) + '\n',
                    lines[5] + '\n' # Preserve previous BOI setting
                ])
         
            # Helper functions (reused)
            def GetCenterPoint(ele_id):
                bBox = doc.GetElement(ele_id).get_BoundingBox(None)
                if bBox:
                    center = (bBox.Max + bBox.Min) / 2
                    return center
                else:
                    return XYZ(0, 0, 0)
         
            def myround(x, multiple):
                return multiple * math.ceil(x / multiple)
         
            def get_reference_level(hanger):
                level_id = hanger.LevelId
                level = doc.GetElement(level_id)
                return level
         
            def get_level_elevation(level):
                if level:
                    return level.Elevation
                return 0.0
         
            # Get bounding box of P-Trap
            ptrap_bb = element.get_BoundingBox(curview)
            if not ptrap_bb:
                print("No bounding box found for P-Trap {}".format(element.Id))
                raise Exception("P-Trap bounding box not found.")
         
            # Calculate bottom middle point (always align to bottom of insulation if present)
            thick = element.InsulationThickness if hasattr(element, 'InsulationThickness') and element.HasInsulation else 0.0
            center_xy = (ptrap_bb.Max + ptrap_bb.Min) / 2
            bottom_z = ptrap_bb.Min.Z - thick if element.HasInsulation else ptrap_bb.Min.Z
            target_pos = XYZ(center_xy.X, center_xy.Y, bottom_z)
         
            # Get connectors C2 and C3 to determine angle
            connector_manager = element.ConnectorManager
            c2 = None
            c3 = None
            for connector in connector_manager.Connectors:
                if connector.Id == 1:
                    c2 = connector
                elif connector.Id == 2:
                    c3 = connector
            if not (c2 and c3):
                print("Connectors C2 and/or C3 not found for P-Trap {}".format(element.Id))
                raise Exception("Required connectors not found.")
         
            # Calculate direction vector between C2 and C3 in XY plane
            c2_pos = c2.Origin
            c3_pos = c3.Origin
            dir_vec = (c3_pos - c2_pos).Normalize()
            dir_vec = XYZ(dir_vec.X, dir_vec.Y, 0).Normalize() # Project to XY plane
            angle_rad = math.atan2(dir_vec.Y, dir_vec.X) + math.pi
         
            # Create hanger
            t = Transaction(doc, 'Place Trapeze Hanger on P-Trap')
            t.Start()
            try:
                hanger = FabricationPart.CreateHanger(doc, fab_btn, 0, level_id)
                if not hanger:
                    print("Failed to create hanger for P-Trap")
                    t.RollBack()
                    raise Exception("Hanger creation failed.")
            except Exception as e:
                print("Error creating hanger: {}".format(str(e)))
                t.RollBack()
                raise Exception("Hanger creation failed.")
         
            # Set hanger dimensions using user-defined width
            for dim in hanger.GetDimensions():
                dim_name = dim.Name
                try:
                    if dim_name == "Width":
                        hanger.SetDimensionValue(dim, (trap_width - 0.166666))
                    if dim_name == "Bearer Extn":
                        hanger.SetDimensionValue(dim, 0.25)
                except Exception as e:
                    print("Error setting dimension '{}' for hanger {}: {}".format(dim_name, hanger.Id, str(e)))
         
            # Rotate hanger
            center = GetCenterPoint(hanger.Id)
            z_axis = Line.CreateBound(center, center + XYZ(0, 0, 1))
            try:
                ElementTransformUtils.RotateElement(doc, hanger.Id, z_axis, angle_rad)
            except Exception as e:
                print("Error rotating hanger {}: {}".format(hanger.Id, str(e)))
                t.RollBack()
                raise Exception("Failed to rotate hanger.")
         
            # Move hanger to target position
            center = GetCenterPoint(hanger.Id)
            translation = target_pos - center
            try:
                ElementTransformUtils.MoveElement(doc, hanger.Id, translation)
            except Exception as e:
                print("Error moving hanger {}: {}".format(hanger.Id, str(e)))
                t.RollBack()
                raise Exception("Failed to move hanger.")
         
            # Set offset
            reference_level = get_reference_level(hanger)
            elevation = get_level_elevation(reference_level)
            try:
                offset_param = hanger.get_Parameter(BuiltInParameter.FABRICATION_OFFSET_PARAM)
                offset_value = bottom_z - elevation
                offset_param.Set(offset_value)
            except Exception as e:
                print("Error setting offset for hanger {}: {}".format(hanger.Id, str(e)))
         
            # Attach to structure if selected
            if AtoS:
                try:
                    hanger.GetRodInfo().AttachToStructure()
                except Exception as e:
                    print("Error attaching hanger {} to structure: {}".format(hanger.Id, str(e)))
         
            t.Commit()
     
        else:
            pass
except Exception as e:
    TaskDialog.Show("Error", "Script error: {}".format(str(e)))