
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
    if 'PLUMBING: DOMESTIC COLD WATER' in servicenamelist:
        components = [
            Label('Choose Service to Place Trapeze on:'),
            ComboBox('ServiceName', servicenamelist, sort=False, default= lines[3]),
            Button('Ok')
            ]
        form = FlexForm('Fabrication Service', components)
        form.show()
    else:
        components = [
            Label('Choose Service to Place Trapeze on:'),
            ComboBox('ServiceName', servicenamelist, sort=False),
            Button('Ok')
            ]
        form = FlexForm('Fabrication Service', components)
        form.show()
    # Convert dialog input into variable
    SelectedServiceName = (form.values['ServiceName'])

    # Gets matching index of selected element service from the servicenamelist
    Servicenum = servicenamelist.index(SelectedServiceName)


    servicelist = []    
    servicelist.append(LoadedServices)
    FabricationService = servicelist[0]

    groupindexlist = []
    groupnamelist = []

    # Checks revit version and uses different code for versions newer than 2022
    if RevitINT > 2022:
        try:
            groupindexlist.append(LoadedServices[Servicenum].PaletteCount)
            numrange = range(LoadedServices[Servicenum].PaletteCount)
            for Item3 in numrange:
                try:
                    groupnamelist.append(FabricationService[Servicenum].GetPaletteName(Item3))
                except:
                    groupnamelist.append(FabricationService[Servicenum].GetPaletteName(Item3))
        except:
            groupindexlist.append(LoadedServices[Servicenum].PaletteCount)
            numrange = range(LoadedServices[Servicenum].PaletteCount)
            for Item3 in numrange:
                try:
                    groupnamelist.append(FabricationService[Servicenum].GetPaletteName(Item3))
                except:
                    groupnamelist.append(FabricationService[Servicenum].GetPaletteName(Item3))
    else:
        try:
            groupindexlist.append(LoadedServices[Servicenum].GroupCount)
            numrange = range(LoadedServices[Servicenum].GroupCount)
            for Item3 in numrange:
                try:
                    groupnamelist.append(FabricationService[Servicenum].GetGroupName(Item3))
                except:
                    groupnamelist.append(FabricationService[Servicenum].GetGroupName(Item3))
        except:
            groupindexlist.append(LoadedServices[Servicenum].GroupCount)
            numrange = range(LoadedServices[Servicenum].GroupCount)
            for Item3 in numrange:
                try:
                    groupnamelist.append(FabricationService[Servicenum].GetGroupName(Item3))
                except:
                    groupnamelist.append(FabricationService[Servicenum].GetGroupName(Item3))

    if 'SUPPORTS' in groupnamelist:
        SelectedServicegroupname = 'SUPPORTS'
        Servicegroupnum = groupnamelist.index(SelectedServicegroupname)
    else:
        # Display dialog
        components = [
            Label('Choose Service Palette:'),
            ComboBox('Servicegroupnum', groupnamelist, sort=False),
            Button('Ok')
            ]
        form = FlexForm('Group', components)
        form.show()

        # Convert else dialog input into variable
        SelectedServicegroupname = (form.values['Servicegroupnum'])
        Servicegroupnum = groupnamelist.index(SelectedServicegroupname)

    buttoncount = LoadedServices[Servicenum].GetButtonCount(Servicegroupnum)

    buttonnames = []

    count = 0
    while count < buttoncount :
        bt = LoadedServices[Servicenum].GetButton(Servicegroupnum, count)
        count = count + 1
        if bt.IsAHanger:
            buttonnames.append(bt.Name)	

    # Display dialog
    # if lines[0] in buttonnames:
    components = [
        Label('Choose Hanger:'),
        ComboBox('Buttonnum', buttonnames, sort=False, default=lines[0]),
        Label('Distance from End (Ft):'),
        TextBox('EndDist', lines[1]),
        Label('Hanger Spacing (Ft):'),
        TextBox('Spacing', lines[2]),
        CheckBox('checkboxBOI', 'Align Trapeze to Bottom of Insulation', default= checkboxdefBOI),
        CheckBox('checkboxvalue', 'Attach to Structure', default= checkboxdef),
        Button('Ok')
        ]
    form = FlexForm('Hanger and Spacing', components)
    form.show()
    # else:
        # components = [
            # Label('Choose Hanger:'),
            # ComboBox('Buttonnum', buttonnames, sort=False),
            # Label('Distance from End (Ft):'),
            # TextBox('EndDist', lines[1]),
            # Label('Hanger Spacing (Ft):'),
            # TextBox('Spacing', lines[2]),
            # CheckBox('checkboxBOI', 'Align Trapeze to BOI', default= checkboxdefBOI),
            # CheckBox('checkboxvalue', 'Attach to Structure', default= checkboxdef),
            # Button('Ok')
            # ]
        # form = FlexForm('Hanger and Spacing', components)
        # form.show()

    # Convert dialog input into variable
    Selectedbutton = (form.values['Buttonnum'])
    Buttonnum = buttonnames.index(Selectedbutton)
    distancefromend = float(form.values['EndDist'])
    Spacing = float(form.values['Spacing'])
    BOITrap = (form.values['checkboxBOI'])
    AtoS = (form.values['checkboxvalue'])

    # write values to text file for future retrieval
    with open((filepath), 'w') as the_file:
        line1 = (Selectedbutton + '\n')
        line2 = (str(distancefromend) + '\n')
        line3 = (str(Spacing) + '\n')
        line4 = (SelectedServiceName + '\n')
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
        qtyofhgrs = int(delta_x / Spacing)
    if (delta_y) > (delta_x):
        qtyofhgrs = int(delta_y / Spacing)
    
    IncrementSpacing = distancefromend
    #-----------------------------------------------------------------------------------PLACING TRAPS
    for hgr in range(qtyofhgrs):

        t = Transaction(doc, 'Place Trapeze Hanger')
        t.Start()
        #--------------DRAWS TRAP AT 0,0,0--------------#
        hanger = FabricationPart.CreateHanger(doc, FabricationServiceButton, 0, level_id)
        #--------------DRAWS TRAP AT 0,0,0--------------#
        t.Commit()

        t = Transaction(doc, 'Modify Trapeze Hanger')
        t.Start()

        X_side_xyz = XYZ(combined_bounding_box.Min.X + IncrementSpacing, 
                            combined_bounding_box_Center.Y, 
                            combined_bounding_box_Center.Z)
        Y_side_xyz = XYZ(combined_bounding_box_Center.X, 
                            combined_bounding_box.Min.Y + IncrementSpacing, 
                            combined_bounding_box_Center.Z)
        IncrementSpacing = IncrementSpacing + Spacing

    #-----------------------------------------------------------------------------------TRAPS IN X DIRECTION, MOVES AND MODIFIES PLACED TRAPS ABOVE
        if (delta_x) > (delta_y):
            for dim in hanger.GetDimensions():
                Dimensions.append(dim.Name)
                if dim.Name == "Width":
                    width_value = hanger.GetDimensionValue(dim)
                    hanger.SetDimensionValue(dim, delta_y)
                if dim.Name == "Bearer Extn":
                    bearer_value = hanger.GetDimensionValue(dim)
                    hanger.SetDimensionValue(dim, 0.33333)
                if dim.Name == "Width":
                    width_value = hanger.GetDimensionValue(dim)
                if dim.Name == "Bearer Extn":
                    bearer_value = hanger.GetDimensionValue(dim)
                    in_bvalue = (bearer_value * 12)
                    bvalue_abvstd = in_bvalue - 4.0
                    in_wvalue = (width_value * 12)
                    rnd_value = myround((in_bvalue + in_wvalue + bvalue_abvstd), 2)
                    abv_value = rnd_value - in_wvalue
                    hlf_diff = (abv_value - 4.0) / 2
                    new_value = (abv_value - hlf_diff) / 12
                    hanger.SetDimensionValue(dim, new_value)
                translation = X_side_xyz - GetCenterPoint(hanger.Id)
                ElementTransformUtils.MoveElement(doc, hanger.Id, translation)
                if BOITrap:
                    hanger.get_Parameter(BuiltInParameter.FABRICATION_OFFSET_PARAM).Set(PRTElevation)
                else:
                    hanger.get_Parameter(BuiltInParameter.FABRICATION_OFFSET_PARAM).Set(combined_bounding_box.Min.Z)

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

            for dim in hanger.GetDimensions():
                Dimensions.append(dim.Name)
                if dim.Name == "Width":
                    width_value = hanger.GetDimensionValue(dim)
                    hanger.SetDimensionValue(dim, delta_x)
                if dim.Name == "Bearer Extn":
                    bearer_value = hanger.GetDimensionValue(dim)
                    hanger.SetDimensionValue(dim, 0.33333)
                if dim.Name == "Width":
                    width_value = hanger.GetDimensionValue(dim)
                if dim.Name == "Bearer Extn":
                    bearer_value = hanger.GetDimensionValue(dim)
                    in_bvalue = (bearer_value * 12)
                    bvalue_abvstd = in_bvalue - 4.0
                    in_wvalue = (width_value * 12)
                    rnd_value = myround((in_bvalue + in_wvalue + bvalue_abvstd), 2)
                    abv_value = rnd_value - in_wvalue
                    hlf_diff = (abv_value - 4.0) / 2
                    new_value = (abv_value - hlf_diff) / 12
                    hanger.SetDimensionValue(dim, new_value)
                translation = Y_side_xyz - GetCenterPoint(hanger.Id)
                ElementTransformUtils.MoveElement(doc, hanger.Id, translation)
                if BOITrap:
                    hanger.get_Parameter(BuiltInParameter.FABRICATION_OFFSET_PARAM).Set(PRTElevation)
                else:
                    hanger.get_Parameter(BuiltInParameter.FABRICATION_OFFSET_PARAM).Set(combined_bounding_box.Min.Z)
        if AtoS:
            hanger.GetRodInfo().AttachToStructure()
        t.Commit()
else:
    print 'Coming Soon... \nYou will be able to place a trapeze on a ptrap'

