#Imports
import Autodesk
from Autodesk.Revit.DB import Transaction, FabricationConfiguration, BuiltInParameter, FabricationPart, FabricationServiceButton, FabricationService, XYZ
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from rpw.ui.forms import FlexForm, Label, ComboBox, TextBox, Separator, Button, CheckBox
import math
import os

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)

try:
    # selection
    selected_element = uidoc.Selection.PickObject(ObjectType.Element, 'Select a Fabrication Part')
    element = doc.GetElement(selected_element.ElementId)

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
    filepath = os.path.join(folder_name, 'Ribbon_PlaceHangers.txt')

    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
    if not os.path.exists(filepath):
        with open((filepath), 'w') as the_file:
            line1 = (str(buttonnames[0]) + '\n')
            line2 = ('1' + '\n')
            line3 = '4'
            the_file.writelines([line1, line2, line3])

    # read text file for stored values and show them in dialog
    with open((filepath), 'r') as file:
        lines = file.readlines()
        lines = [line.rstrip() for line in lines]

    if len(lines) < 3:
        with open((filepath), 'w') as the_file:
            line1 = (str(buttonnames[0]) + '\n')
            line2 = ('1' + '\n')
            line3 = '4'
            the_file.writelines([line1, line2, line3])

    # read text file for stored values and show them in dialog
    with open((filepath), 'r') as file:
        lines = file.readlines()
        lines = [line.rstrip() for line in lines]

    # Display dialog
    if lines[0] in buttonnames:
        components = [
            Label('Choose Hanger:'),
            ComboBox('Buttonnum', buttonnames, sort=False, default=lines[0]),
            Label('Distance from End (Ft):'),
            TextBox('EndDist', lines[1]),
            Label('Hanger Spacing (Ft):'),
            TextBox('Spacing', lines[2]),
            CheckBox('checkboxvalue', 'Attach to Structure', default=True),
            CheckBox('checkboxjointvalue', 'Support Joints (cannot disable yet...)', default=True),
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
            CheckBox('checkboxvalue', 'Attach to Structure', default=True),
            CheckBox('checkboxjointvalue', 'Support Joints (cannot disable yet...)', default=True),
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
    AtoS = (form.values['checkboxvalue'])
    SupportJoint = (form.values['checkboxjointvalue'])

    # write values to text file for future retrieval
    with open((filepath), 'w') as the_file:
        line1 = (Selectedbutton + '\n')
        line2 = (str(distancefromend) + '\n')
        line3 = str(Spacing)
        the_file.writelines([line1, line2, line3])
        
    # Check if the button selected is valid   
    validbutton = FabricationService[Servicenum].IsValidButtonIndex(Servicegroupnum,Buttonnum)
    FabricationServiceButton = FabricationService[Servicenum].GetButton(Servicegroupnum,Buttonnum)

    # Prompt user to select pipes (with filter)
    class CustomISelectionFilter(ISelectionFilter):
        def __init__(self, nom_categorie):
            self.nom_categorie = nom_categorie
        def AllowElement(self, e):
            if e.LookupParameter('Fabrication Service').AsValueString() == self.nom_categorie:
                return True
            else:
                return False
        def AllowReference(self, ref, point):
            return True

    pipesel = uidoc.Selection.PickObjects(ObjectType.Element,
    CustomISelectionFilter(parameters), "Select pipes to place hangers on")            
    Pipe = [doc.GetElement( elId ) for elId in pipesel]

    # start a transaction to modify model
    t = Transaction(doc, 'Place Hangers')
    t.Start()

    for e in Pipe:
        if e.LookupParameter('Part Pattern Number').AsInteger() in (2041, 866, 40):
            # get length of pipe
            pipelen = e.CenterlineLength
            # test if pipe is long enough for hanger
            if pipelen > distancefromend:
                # block of code to get connectors on both ends of pipe
                pipe_connector = e.ConnectorManager.Connectors
                # checking if support joint option was selected in dialog
                if SupportJoint == True:
                    # setting up and increment addition loop
                    IncrementSpacing = distancefromend
                    # block of code to get connectors on both ends of pipe
                    try:
                        for connector in pipe_connector:
                            # adding hangers to each end of pipe by end distance specified
                            FabricationPart.CreateHanger(doc, FabricationServiceButton, e.Id, connector, distancefromend, AtoS)
                        # testing if pipe is long enough for hangers and spacing
                        if pipelen > (int(Spacing) + (distancefromend * 2)):
                            
                            # calculating how many hangers spaced on pipe are required
                            qtyofhgrs = range(int((math.floor(pipelen) - (distancefromend * 3)) / Spacing))
                            
                            # looping thru qty of hangers and placing them
                            for hgr in qtyofhgrs:
                                IncrementSpacing = IncrementSpacing + Spacing
                                FabricationPart.CreateHanger(doc, FabricationServiceButton, e.Id, connector, IncrementSpacing, AtoS)
                    except:
                        pass

    if SupportJoint == False:
        # Get total run length and create ordered list of pipe segments
        total_run_length = 0
        pipe_segments = []
        
        for pipe in Pipe:
            if pipe.LookupParameter('Part Pattern Number').AsInteger() in (2041, 866, 40):
                # Create tuple of (pipe, start_position)
                pipe_segments.append((pipe, total_run_length))
                total_run_length += pipe.CenterlineLength

        # Calculate number of hangers needed for entire run
        first_hanger_pos = distancefromend
        last_hanger_pos = total_run_length - distancefromend
        
        if last_hanger_pos > first_hanger_pos:
            current_position = first_hanger_pos
            
            while current_position <= last_hanger_pos:
                # Find which pipe segment this position falls on
                current_segment = None
                local_position = current_position
                
                for pipe, start_pos in pipe_segments:
                    pipe_end = start_pos + pipe.CenterlineLength
                    if start_pos <= current_position < pipe_end:
                        current_segment = pipe
                        local_position = current_position - start_pos
                        break
                
                if current_segment:
                    try:
                        # Get first connector of current pipe segment
                        pipe_connector = current_segment.ConnectorManager.Connectors
                        first_connector = next(iter(pipe_connector))
                        
                        # Place hanger at calculated position
                        FabricationPart.CreateHanger(
                            doc, 
                            FabricationServiceButton, 
                            current_segment.Id, 
                            first_connector, 
                            local_position,
                            AtoS
                        )
                    except Exception as e:
                        print("Failed to place hanger")
                
                # Move to next hanger position
                current_position += Spacing

    # end transaction
    t.Commit()
except:
    pass