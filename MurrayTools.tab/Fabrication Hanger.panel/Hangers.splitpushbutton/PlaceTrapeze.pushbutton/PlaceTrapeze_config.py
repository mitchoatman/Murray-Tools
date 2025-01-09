#------------------------------------------------------------------------------------IMPORTS
import Autodesk
from Autodesk.Revit.DB import Transaction, FabricationConfiguration, BuiltInParameter, FabricationPart, FabricationServiceButton, \
                                FabricationService, XYZ, ElementTransformUtils, BoundingBoxXYZ, Transform, Line
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from rpw.ui.forms import FlexForm, Label, ComboBox, TextBox, Separator, Button, CheckBox
import math
import os

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
        ServiceBtns = Config.GetAllLoadedItemFiles()

        for Item1 in LoadedServices:
            try:
                servicenamelist.append(Item1.Name)
            except:
                servicenamelist.append([])

        # Gets matching index of selected element service from the servicenamelist
        Servicenum = servicenamelist.index(parameters)

        servicelist = []    
        servicelist.append(LoadedServices)
        FabricationService = servicelist[0]

        # Find all hanger buttons across all palettes/groups
        buttonnames = []
        button_data = []  # Store tuple of (palette_idx, button_idx, button_name)

        if RevitINT > 2022:
            palette_count = LoadedServices[Servicenum].PaletteCount
            for palette_idx in range(palette_count):
                buttoncount = LoadedServices[Servicenum].GetButtonCount(palette_idx)
                for btn_idx in range(buttoncount):
                    bt = LoadedServices[Servicenum].GetButton(palette_idx, btn_idx)
                    if bt.IsAHanger:
                        buttonnames.append(bt.Name)
                        button_data.append((palette_idx, btn_idx, bt.Name))
        else:
            group_count = LoadedServices[Servicenum].GroupCount
            for group_idx in range(group_count):
                buttoncount = LoadedServices[Servicenum].GetButtonCount(group_idx)
                for btn_idx in range(buttoncount):
                    bt = LoadedServices[Servicenum].GetButton(group_idx, btn_idx)
                    if bt.IsAHanger:
                        buttonnames.append(bt.Name)
                        button_data.append((group_idx, btn_idx, bt.Name))

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

        # Display dialog
        if lines[0] in buttonnames:
            components = [
                Label('Choose Hanger:'),
                ComboBox('Buttonnum', buttonnames, sort=False, default=lines[0]),
                Label('Distance from End (Ft):'),
                TextBox('EndDist', lines[1]),
                Label('Hanger Spacing (Ft):'),
                TextBox('Spacing', lines[2]),
                CheckBox('checkboxBOI', 'Align Trapeze to Bottom of Insulation', default=checkboxdefBOI),
                CheckBox('checkboxvalue', 'Attach to Structure', default=checkboxdef),
                Button('Ok')
            ]
            form = FlexForm('Hanger and Spacing', components)
            form.show()
        else:
            components = [
                Label('Choose Hanger:'),
                ComboBox('Buttonnum', buttonnames, sort=False),
                Label('Distance from End (Ft):'),
                TextBox('EndDist', lines[1]),
                Label('Hanger Spacing (Ft):'),
                TextBox('Spacing', lines[2]),
                CheckBox('checkboxBOI', 'Align Trapeze to BOI', default=checkboxdefBOI),
                CheckBox('checkboxvalue', 'Attach to Structure', default=checkboxdef),
                Button('Ok')
            ]
            form = FlexForm('Hanger and Spacing', components)
            form.show()

        # Convert dialog input into variable
        Selectedbutton = (form.values['Buttonnum'])
        # Find the palette_idx and button_idx for the selected button
        for palette_idx, btn_idx, btn_name in button_data:
            if btn_name == Selectedbutton:
                Servicegroupnum = palette_idx
                Buttonnum = btn_idx
                break

        distancefromend = float(form.values['EndDist'])
        Spacing = float(form.values['Spacing'])
        BOITrap = (form.values['checkboxBOI'])
        AtoS = (form.values['checkboxvalue'])

        # write values to text file for future retrieval
        with open((filepath), 'w') as the_file:
            line1 = (Selectedbutton + '\n')
            line2 = (str(distancefromend) + '\n')
            line3 = (str(Spacing) + '\n')
            line4 = (parameters + '\n')
            line5 = (str(AtoS) + '\n')
            line6 = str(BOITrap)
            the_file.writelines([line1, line2, line3, line4, line5, line6])
            
        # Check if the button selected is valid   
        validbutton = FabricationService[Servicenum].IsValidButtonIndex(Servicegroupnum,Buttonnum)
        FabricationServiceButton = FabricationService[Servicenum].GetButton(Servicegroupnum,Buttonnum)

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

        X_side_xyz = XYZ(combined_bounding_box.Min.X + distancefromend, 
                            combined_bounding_box_Center.Y, 
                            combined_bounding_box_Center.Z)
        Y_side_xyz = XYZ(combined_bounding_box_Center.X, 
                            combined_bounding_box.Min.Y + distancefromend, 
                            combined_bounding_box_Center.Z)

        X_side_xyz_opp = XYZ(combined_bounding_box.Max.X - distancefromend, 
                            combined_bounding_box_Center.Y, 
                            combined_bounding_box_Center.Z)
        Y_side_xyz_opp = XYZ(combined_bounding_box_Center.X, 
                            combined_bounding_box.Max.Y - distancefromend, 
                            combined_bounding_box_Center.Z)

        delta_x = abs(combined_bounding_box.Max.X - combined_bounding_box.Min.X)
        delta_y = abs(combined_bounding_box.Max.Y - combined_bounding_box.Min.Y)

        #-----------------------------------------------------------------------------------SETUP SPACING
        Dimensions = []

        # Calculate how many hangers in the run
        if (delta_x) > (delta_y):
            qtyofhgrs = int(math.ceil(delta_x / Spacing))
        if (delta_y) > (delta_x):
            qtyofhgrs = int(math.ceil(delta_y / Spacing))
        
        IncrementSpacing = distancefromend
        #-----------------------------------------------------------------------------------PLACING TRAPS
        
        hangers = []
        
        t = Transaction(doc, 'Place Trapeze Hanger')
        t.Start()
        for hgr in range(qtyofhgrs):
            #--------------DRAWS TRAP AT 0,0,0--------------#
            hanger = FabricationPart.CreateHanger(doc, FabricationServiceButton, 0, level_id)
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
            IncrementSpacing = IncrementSpacing + Spacing

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
        print 'Coming Soon... \nYou will be able to place a trapeze on a ptrap'
except:
    pass