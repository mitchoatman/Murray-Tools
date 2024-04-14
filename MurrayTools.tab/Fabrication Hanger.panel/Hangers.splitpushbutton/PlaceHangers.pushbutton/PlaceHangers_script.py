
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
RevitINT = float (RevitVersion)

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
    Buttonnum = buttonnames.index(Selectedbutton)
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
            return true

    pipesel = uidoc.Selection.PickObjects(ObjectType.Element,
    CustomISelectionFilter(parameters), "Select pipes to place hangers on")            
    Pipe = [doc.GetElement( elId ) for elId in pipesel]

    # start a transaction to modify model
    t = Transaction(doc, 'Place Hangers')
    t.Start()

    for e in Pipe:
        if e.LookupParameter('Part Pattern Number').AsInteger() in (2041, 866, 40):
            # get length of pipe
            pipelen = e.LookupParameter('Length').AsDouble()
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
                            #print connector.Origin.Z
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
                    for connector in pipe_connector:
                        connector = connector
                    # setting up and increment addition loop
                    IncrementSpacing = distancefromend
                    try:
                        # adding hanger to one end of pipe by end distance specified
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
    # end transaction
    t.Commit()
except:
    pass

