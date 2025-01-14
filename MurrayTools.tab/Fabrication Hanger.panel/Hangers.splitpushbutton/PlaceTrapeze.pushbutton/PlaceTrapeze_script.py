#------------------------------------------------------------------------------------IMPORTS
import Autodesk
from Autodesk.Revit.DB import Transaction, FabricationConfiguration, BuiltInParameter, FabricationPart, FabricationServiceButton, \
                                FabricationService, XYZ, ElementTransformUtils, BoundingBoxXYZ, Transform, Line
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import math
import os

import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
clr.AddReference("System")

from System.Windows.Forms import *
from System.Drawing import Point, Size, Font
from System import Array
#------------------------------------------------------------------------------------DEFINE SOME VARIABLES EASY USE
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float (RevitVersion)

#------------------------------------------------------------------------------------SELECTING ELEMENTS
try:
    selected_element = uidoc.Selection.PickObject(ObjectType.Element, 'Select OUTSIDE Pipe')
    if doc.GetElement(selected_element.ElementId).ItemCustomId != 916:
        selected_element1 = uidoc.Selection.PickObject(ObjectType.Element, 'Select OPPOSITE OUTSIDE Pipe')
        element = doc.GetElement(selected_element.ElementId)
        element1 = doc.GetElement(selected_element1.ElementId)
        selected_elements = [element, element1]

        level_id = element.LevelId

        #FUNCTION TO GET PARAMETER VALUE  change "AsDouble()" to "AsString()" to change data type.
        def get_parameter_value(element, parameterName):
            return element.LookupParameter(parameterName).AsDouble()

        # Gets bottom elevation of selected pipe
        if element and RevitINT > 2022:
            PRTElevation = get_parameter_value(element, 'Lower End Bottom Elevation')

        if element and RevitINT < 2023:
            PRTElevation = get_parameter_value(element, 'Bottom')

        # Gets servicename of selection
        parameters = element.LookupParameter('Fabrication Service').AsValueString()

        servicenamelist = []
        Config = FabricationConfiguration.GetFabricationConfiguration(doc)
        LoadedServices = Config.GetAllLoadedServices()

        for Item1 in LoadedServices:
            try:
                servicenamelist.append(Item1.Name)
            except:
                servicenamelist.append([])

        # Gets matching index of selected element service from the servicenamelist
        Servicenum = servicenamelist.index(parameters)

        # Find all hanger buttons across all palettes/groups
        buttonnames = []
        button_data = []  # Store tuple of (palette_idx, button_idx, button_name)

        buttonnames = []
        button_data = []

        buttonnames = []
        unique_hangers = set()  # To track unique hanger button names

        for service_idx, service in enumerate(LoadedServices):  # Loop through all loaded services
            if RevitINT > 2022:
                palette_count = service.PaletteCount  # Get palette count for Revit 2023 and newer
            else:
                palette_count = service.GroupCount  # Get group count for older Revit versions

            for palette_idx in range(palette_count):  # Loop through each palette/group in the service
                buttoncount = service.GetButtonCount(palette_idx)  # Get button count for the current palette/group

                for btn_idx in range(buttoncount):  # Loop through each button in the palette/group
                    bt = service.GetButton(palette_idx, btn_idx)  # Get the button

                    if bt.IsAHanger and bt.Name not in unique_hangers:  # Check if the button is a hanger and is unique
                        unique_hangers.add(bt.Name)  # Add to the set of unique hangers
                        buttonnames.append(bt.Name)  # Add to the final list of button names

        folder_name = "c:\\Temp"
        filepath = os.path.join(folder_name, 'Ribbon_PlaceTrapeze.txt')
        if not os.path.exists(folder_name):
            os.makedirs(folder_name)
        if not os.path.exists(filepath):
            with open((filepath), 'w') as the_file:
                line1 = ('1.625 Single Strut Trapeze' + '\n')
                line2 = ('1.0' + '\n')
                line3 = ('8.0' + '\n')
                line4 = ('PLUMBING: DOMESTIC COLD WATER' + '\n')
                line5 = ('True'+ '\n')
                line6 = 'True'
                the_file.writelines([line1, line2, line3, line4, line5, line6])

        # read text file for stored values and show them in dialog
        with open((filepath), 'r') as file:
            lines = file.readlines()
            lines = [line.rstrip() for line in lines]

        if len(lines) < 6:
            with open((filepath), 'w') as the_file:
                line1 = ('1.625 Single Strut Trapeze' + '\n')
                line2 = ('1.0' + '\n')
                line3 = ('8.0' + '\n')
                line4 = ('PLUMBING: DOMESTIC COLD WATER' + '\n')
                line5 = ('True'+ '\n')
                line6 = 'True'
                the_file.writelines([line1, line2, line3, line4, line5, line6])

        # read text file for stored values and show them in dialog
        with open((filepath), 'r') as file:
            lines = file.readlines()
            lines = [line.rstrip() for line in lines]

        if lines[4] == 'False':
            checkboxdef = False
        else:
            checkboxdef = True

        if lines[5] == 'False':
            checkboxdefBOI = False
        else:
            checkboxdefBOI = True

    #-----------------------------------------------------------

        class HangerSpacingDialog(Form):
            def __init__(self, buttonnames, lines, checkboxdefBOI, checkboxdef):
                self.Text = "Hanger and Spacing"
                self.Size = Size(350, 360)
                self.StartPosition = FormStartPosition.CenterScreen
                self.FormBorderStyle = FormBorderStyle.FixedDialog

                # Choose Hanger Label
                label_hanger = Label()
                label_hanger.Text = "Choose Hanger:"
                label_hanger.Location = Point(10, 10)
                label_hanger.Size = Size(300, 20)
                label_hanger.Font = Font("Arial", 10)
                self.Controls.Add(label_hanger)

                # Hanger ComboBox
                self.combobox_hanger = ComboBox()
                self.combobox_hanger.Location = Point(10, 31)
                self.combobox_hanger.Size = Size(300, 20)
                self.combobox_hanger.DropDownStyle = ComboBoxStyle.DropDownList
                # Directly pass the list as an array
                self.combobox_hanger.Items.AddRange(Array[object](buttonnames))
                if lines[0] in buttonnames:
                    self.combobox_hanger.SelectedItem = lines[0]
                self.Controls.Add(self.combobox_hanger)

                # Distance from End Label
                label_end_dist = Label()
                label_end_dist.Text = "Distance from End (Ft):"
                label_end_dist.Font = Font("Arial", 10)
                label_end_dist.Location = Point(10, 60)
                label_end_dist.Size = Size(300, 20)
                self.Controls.Add(label_end_dist)

                # Distance from End TextBox
                self.textbox_end_dist = TextBox()
                self.textbox_end_dist.Location = Point(10, 80)
                self.textbox_end_dist.Text = lines[1]
                self.Controls.Add(self.textbox_end_dist)

                # Hanger Spacing Label
                label_spacing = Label()
                label_spacing.Text = "Hanger Spacing (Ft):"
                label_spacing.Font = Font("Arial", 10)
                label_spacing.Location = Point(10, 110)
                label_spacing.Size = Size(300, 20)
                self.Controls.Add(label_spacing)

                # Hanger Spacing TextBox
                self.textbox_spacing = TextBox()
                self.textbox_spacing.Location = Point(10, 130)
                self.textbox_spacing.Text = lines[2]
                self.Controls.Add(self.textbox_spacing)

                # Align Trapeze CheckBox
                self.checkbox_boi = CheckBox()
                self.checkbox_boi.Text = "Align Trapeze to Bottom of Insulation"
                self.checkbox_boi.Font = Font("Arial", 10)
                self.checkbox_boi.Location = Point(10, 160)
                self.checkbox_boi.Size = Size(300, 20)
                self.checkbox_boi.Checked = checkboxdefBOI
                self.Controls.Add(self.checkbox_boi)

                # Attach to Structure CheckBox
                self.checkbox_attach = CheckBox()
                self.checkbox_attach.Text = "Attach to Structure"
                self.checkbox_attach.Font = Font("Arial", 10)
                self.checkbox_attach.Location = Point(10, 190)
                self.checkbox_attach.Size = Size(300, 20)
                self.checkbox_attach.Checked = checkboxdef
                self.Controls.Add(self.checkbox_attach)

                # Choose Service Label
                label_service = Label()
                label_service.Text = "Choose Service to Draw Hanger on:"
                label_service.Location = Point(10, 220)
                label_service.Size = Size(300, 20)
                label_service.Font = Font("Arial", 10)
                self.Controls.Add(label_service)

                # Service ComboBox
                self.combobox_service = ComboBox()
                self.combobox_service.Location = Point(10, 240)
                self.combobox_service.Size = Size(300, 20)
                self.combobox_service.DropDownStyle = ComboBoxStyle.DropDownList
                # Directly pass the list as an array
                self.combobox_service.Items.AddRange(Array[object](servicenamelist))
                if lines[3] in servicenamelist:
                    self.combobox_service.SelectedItem = lines[3]
                self.Controls.Add(self.combobox_service)

                # OK Button
                self.button_ok = Button()
                self.button_ok.Text = "OK"
                self.button_ok.Font = Font("Arial", 10)
                self.button_ok.Location = Point(((self.Width // 2) - 50), 280)
                self.button_ok.Click += self.ok_button_clicked
                self.Controls.Add(self.button_ok)

            def ok_button_clicked(self, sender, event):
                self.DialogResult = DialogResult.OK
                self.Close()


        form = HangerSpacingDialog(buttonnames, lines, checkboxdefBOI, checkboxdef)
        if form.ShowDialog() == DialogResult.OK:
            Selectedbutton = form.combobox_hanger.Text
            distancefromend = form.textbox_end_dist.Text
            Spacing = form.textbox_spacing.Text
            BOITrap = form.checkbox_boi.Checked
            AtoS = form.checkbox_attach.Checked
            SelectedServiceName = form.combobox_service.Text
            
            # Gets matching index of selected service from the servicenamelist
            Servicenum = servicenamelist.index(SelectedServiceName)
            
            # Initialize a flag to track if the button is found
            button_found = False

            # Loop through all services to find the selected button
            for servicenum, service in enumerate(LoadedServices):
                if service.Name == SelectedServiceName:
                    palette_count = service.PaletteCount if RevitINT > 2022 else service.GroupCount
                    for palette_idx in range(palette_count):
                        button_count = service.GetButtonCount(palette_idx)
                        for btn_idx in range(button_count):
                            bt = service.GetButton(palette_idx, btn_idx)
                            if bt.Name == Selectedbutton:  # Match selected button name
                                fab_btn = bt  # Optional: Store the button object for later use
                                button_found = True  # Mark button as found
                                break
                        else:
                            continue  # Continue outer loop if inner loop was not broken
                        break
                    else:
                        continue  # Continue outer loop if inner loop was not broken
                    break

            # Check if the button was found
            if button_found:
                
                # write values to text file for future retrieval
                with open((filepath), 'w') as the_file:
                    line1 = (str(Selectedbutton) + '\n')
                    line2 = (str(distancefromend) + '\n')
                    line3 = (str(Spacing) + '\n')
                    line4 = (SelectedServiceName + '\n')
                    line5 = (str(AtoS) + '\n')
                    line6 = str(BOITrap)
                    the_file.writelines([line1, line2, line3, line4, line5, line6])
                    
                def GetCenterPoint(ele):
                    bBox = doc.GetElement(ele).get_BoundingBox(None)
                    center = (bBox.Max + bBox.Min) / 2
                    return center

                def myround(x, multiple):
                    return multiple * math.ceil(x/multiple)

                first_pipe_bounding_box = element.get_BoundingBox(curview)

                # Determine Rack Direction
                # Calculate the differences (deltas) in X, Y, and Z coordinates
                delta_x = abs(first_pipe_bounding_box.Max.X - first_pipe_bounding_box.Min.X)
                delta_y = abs(first_pipe_bounding_box.Max.Y - first_pipe_bounding_box.Min.Y)

                #-----------------------------------------------------------------------------------WITH INSULATION
                # Initialize variables for the combined bounding box with the coordinates of the first bounding box
                combined_min = first_pipe_bounding_box.Min
                combined_max = first_pipe_bounding_box.Max

                if (delta_x) > (delta_y):
                    # Iterate through the selected pipes to calculate individual bounding boxes
                    for pipe in selected_elements:
                        pipe_bounding_box = pipe.get_BoundingBox(curview)

                        if pipe.HasInsulation:
                            # Update the combined_min and combined_max coordinates
                            pipe_bounding_box.Min = XYZ(pipe_bounding_box.Min.X,
                                                        pipe_bounding_box.Min.Y - pipe.InsulationThickness,
                                                        pipe_bounding_box.Min.Z)
                            
                            pipe_bounding_box.Max = XYZ(pipe_bounding_box.Max.X,
                                                        pipe_bounding_box.Max.Y + pipe.InsulationThickness,
                                                        pipe_bounding_box.Max.Z)

                        # Update the combined_min and combined_max coordinates
                        combined_min = XYZ(min(combined_min.X, pipe_bounding_box.Min.X),
                                            min(combined_min.Y, pipe_bounding_box.Min.Y),
                                            min(combined_min.Z, pipe_bounding_box.Min.Z))

                        combined_max = XYZ(max(combined_max.X, pipe_bounding_box.Max.X),
                                            max(combined_max.Y, pipe_bounding_box.Max.Y),
                                            max(combined_max.Z, pipe_bounding_box.Max.Z))

                if (delta_y) > (delta_x):
                    # Iterate through the selected pipes to calculate individual bounding boxes
                    for pipe in selected_elements:
                        pipe_bounding_box = pipe.get_BoundingBox(curview)

                        if pipe.HasInsulation:
                            # Update the combined_min and combined_max coordinates
                            pipe_bounding_box.Min = XYZ(pipe_bounding_box.Min.X - pipe.InsulationThickness,
                                                        pipe_bounding_box.Min.Y,
                                                        pipe_bounding_box.Min.Z)
                            
                            pipe_bounding_box.Max = XYZ(pipe_bounding_box.Max.X + pipe.InsulationThickness,
                                                        pipe_bounding_box.Max.Y,
                                                        pipe_bounding_box.Max.Z)

                        # Update the combined_min and combined_max coordinates
                        combined_min = XYZ(min(combined_min.X, pipe_bounding_box.Min.X),
                                            min(combined_min.Y, pipe_bounding_box.Min.Y),
                                            min(combined_min.Z, pipe_bounding_box.Min.Z))

                        combined_max = XYZ(max(combined_max.X, pipe_bounding_box.Max.X),
                                            max(combined_max.Y, pipe_bounding_box.Max.Y),
                                            max(combined_max.Z, pipe_bounding_box.Max.Z))

                # Function to get the reference level of a hanger
                def get_reference_level(hanger):
                    level_id = hanger.LevelId
                    level = doc.GetElement(level_id)
                    return level

                # Function to get the elevation of the reference level
                def get_level_elevation(level):
                    if level:
                        return level.Elevation
                    else:
                        return None

                # Create a new combined bounding box using the calculated coordinates
                combined_bounding_box = BoundingBoxXYZ()
                combined_bounding_box.Min = combined_min
                combined_bounding_box.Max = combined_max
                combined_bounding_box_Center = (combined_bounding_box.Max + combined_bounding_box.Min) / 2

                X_side_xyz = XYZ(combined_bounding_box.Min.X + float(distancefromend), 
                                    combined_bounding_box_Center.Y, 
                                    combined_bounding_box_Center.Z)
                Y_side_xyz = XYZ(combined_bounding_box_Center.X, 
                                    combined_bounding_box.Min.Y + float(distancefromend), 
                                    combined_bounding_box_Center.Z)

                X_side_xyz_opp = XYZ(combined_bounding_box.Max.X - float(distancefromend), 
                                    combined_bounding_box_Center.Y, 
                                    combined_bounding_box_Center.Z)
                Y_side_xyz_opp = XYZ(combined_bounding_box_Center.X, 
                                    combined_bounding_box.Max.Y - float(distancefromend), 
                                    combined_bounding_box_Center.Z)

                delta_x = abs(combined_bounding_box.Max.X - combined_bounding_box.Min.X)
                delta_y = abs(combined_bounding_box.Max.Y - combined_bounding_box.Min.Y)

                #-----------------------------------------------------------------------------------SETUP SPACING
                Dimensions = []

                # Calculate how many hangers in the run
                if (delta_x) > (delta_y):
                    qtyofhgrs = int(math.ceil(delta_x / float(Spacing)))
                if (delta_y) > (delta_x):
                    qtyofhgrs = int(math.ceil(delta_y / float(Spacing)))
                
                IncrementSpacing = float(distancefromend)
                #-----------------------------------------------------------------------------------PLACING TRAPS
                
                hangers = []
                
                t = Transaction(doc, 'Place Trapeze Hanger')
                t.Start()
                for hgr in range(qtyofhgrs):
                    #--------------DRAWS TRAP AT 0,0,0--------------#
                    hanger = FabricationPart.CreateHanger(doc, fab_btn, 0, level_id)
                    #--------------DRAWS TRAP AT 0,0,0--------------#

                    # Append each instance to the list
                    hangers.append(hanger)
                t.Commit()

                t = Transaction(doc, 'Modify Trapeze Hanger')
                t.Start()

                for hanger in hangers:
                    X_side_xyz = XYZ(combined_bounding_box.Min.X + IncrementSpacing, 
                                        combined_bounding_box_Center.Y, 
                                        combined_bounding_box_Center.Z)
                    Y_side_xyz = XYZ(combined_bounding_box_Center.X, 
                                        combined_bounding_box.Min.Y + IncrementSpacing, 
                                        combined_bounding_box_Center.Z)
                    IncrementSpacing = IncrementSpacing + float(Spacing)

                #-----------------------------------------------------------------------------------TRAPS IN X DIRECTION, MOVES AND MODIFIES PLACED TRAPS ABOVE
                    if (delta_x) > (delta_y):
                        newwidth = (myround((delta_y * 12), 2) / 12)
                        for dim in hanger.GetDimensions():
                            Dimensions.append(dim.Name)
                            if dim.Name == "Width":
                                width_value = hanger.GetDimensionValue(dim)
                                hanger.SetDimensionValue(dim, delta_y)
                            if dim.Name == "Bearer Extn":
                                bearer_value = hanger.GetDimensionValue(dim)
                                hanger.SetDimensionValue(dim, 0.25)
                            translation = X_side_xyz - GetCenterPoint(hanger.Id)
                            ElementTransformUtils.MoveElement(doc, hanger.Id, translation)
                            reference_level = get_reference_level(hanger)
                            if reference_level:
                                elevation = get_level_elevation(reference_level)
                            if BOITrap:
                                hanger.get_Parameter(BuiltInParameter.FABRICATION_OFFSET_PARAM).Set(PRTElevation)
                            else:
                                hanger.get_Parameter(BuiltInParameter.FABRICATION_OFFSET_PARAM).Set(combined_bounding_box.Min.Z - elevation)

                #-----------------------------------------------------------------------------------TRAPS IN Y DIRECTION, MOVES AND MODIFIES PLACED TRAPS ABOVE
                    if (delta_y) > (delta_x):
                        #---------------ROTATION OF TRAP---------------#
                        # Specify the Z-axis direction (adjust as needed)
                        z_axis_direction = XYZ(0, 0, 1)  # Assuming positive Z direction

                        # Create a list of points for the curve
                        curve_points = [GetCenterPoint(hanger.Id), GetCenterPoint(hanger.Id) + z_axis_direction * 2]  # Adjust the length as needed

                        # Create a curve using the points
                        curve = Autodesk.Revit.DB.Line.CreateBound(curve_points[0], curve_points[1])
                        ElementTransformUtils.RotateElement(doc, hanger.Id, curve, (90.0 * (math.pi / 180.0)))
                        #---------------ROTATION OF TRAP---------------#

                        newwidth = (myround((delta_x * 12), 2) / 12)
                        for dim in hanger.GetDimensions():
                            Dimensions.append(dim.Name)
                            if dim.Name == "Width":
                                width_value = hanger.GetDimensionValue(dim)
                                hanger.SetDimensionValue(dim, delta_x)
                            if dim.Name == "Bearer Extn":
                                bearer_value = hanger.GetDimensionValue(dim)
                                hanger.SetDimensionValue(dim, 0.25)
                            translation = Y_side_xyz - GetCenterPoint(hanger.Id)
                            ElementTransformUtils.MoveElement(doc, hanger.Id, translation)
                            reference_level = get_reference_level(hanger)
                            if reference_level:
                                elevation = get_level_elevation(reference_level)
                            if BOITrap:
                                hanger.get_Parameter(BuiltInParameter.FABRICATION_OFFSET_PARAM).Set(PRTElevation)
                            else:
                                hanger.get_Parameter(BuiltInParameter.FABRICATION_OFFSET_PARAM).Set(combined_bounding_box.Min.Z - elevation)
                    if AtoS:
                        hanger.GetRodInfo().AttachToStructure()
                t.Commit()
            else:
                # If the button was not found, print the error message
                print "'{}' not found in '{}'".format(Selectedbutton, SelectedServiceName)
    else:
        print 'Coming Soon... \nYou will be able to place a trapeze on a ptrap'
except:
    pass