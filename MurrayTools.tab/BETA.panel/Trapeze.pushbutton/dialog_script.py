#Imports
import Autodesk
from Autodesk.Revit.DB import Transaction, FabricationConfiguration, BuiltInParameter, FabricationPart, FabricationServiceButton, FabricationService, XYZ, ElementTransformUtils, BoundingBoxXYZ, Transform, Line
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

def GetCenterPoint(ele):
    bBox = doc.GetElement(ele).get_BoundingBox(None)
    center = (bBox.Max + bBox.Min) / 2
    return center
def myround(x, multiple):
    return multiple * math.ceil(x/multiple)
def get_parameter_value(element, parameterName):
    return element.LookupParameter(parameterName).AsDouble()

# selection
selected_element = uidoc.Selection.PickObject(ObjectType.Element, 'Select a Fabrication Part')
element = doc.GetElement(selected_element.ElementId)
level_id = element.LevelId
#RCKelevation = get_parameter_value(element, 'Lower End Bottom Elevation')


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
    # Display dialog
    components = [
        Label('Choose Service Palette:'),
        ComboBox('Servicegroupnum', groupnamelist, sort=False, default='SUPPORTS'),
        Button('Ok')
        ]
    form = FlexForm('Group', components)
    form.show()
else:
    # Display dialog
    components = [
        Label('Choose Service Palette:'),
        ComboBox('Servicegroupnum', groupnamelist, sort=False),
        Button('Ok')
        ]
    form = FlexForm('Group', components)
    form.show()

# Convert dialog input into variable
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

pipesel = uidoc.Selection.PickObjects(ObjectType.Element, "Select pipes to place hangers on")            
Pipe = [doc.GetElement( elId ) for elId in pipesel]

# Initialize variables for the combined bounding box with the coordinates of the first bounding box
first_pipe_bounding_box = Pipe[0].get_BoundingBox(curview)
combined_min = first_pipe_bounding_box.Min
combined_max = first_pipe_bounding_box.Max

# Iterate through the selected pipes to calculate individual bounding boxes
for pipe in Pipe[1:]:
    # Get the bounding box of the current pipe
    pipe_bounding_box = pipe.get_BoundingBox(curview)

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

# Calculate the differences (deltas) in X, Y, and Z coordinates
delta_x = combined_bounding_box.Max.X - combined_bounding_box.Min.X
delta_y = combined_bounding_box.Max.Y - combined_bounding_box.Min.Y
delta_z = combined_bounding_box.Max.Z - combined_bounding_box.Min.Z

bottom_point = combined_min.Z

# # Calculate the angle along the X-axis
# angle_x_rad = math.atan2(0, delta_x)  # 0 is used for the Y component to measure along the X-axis
# angle_x_deg = math.degrees(angle_x_rad)

# # Calculate the angle along the Y-axis
# angle_y_rad = math.atan2(delta_y, 0)  # 0 is used for the X component to measure along the Y-axis
# angle_y_deg = math.degrees(angle_y_rad)

Dimensions = []

t = Transaction(doc, 'Place Hangers')
t.Start()

#FabricationPart.CreateHanger(doc, FabricationServiceButton, ElementId.Id, Connector, double, bool)
hanger = FabricationPart.CreateHanger(doc, FabricationServiceButton, 0, level_id)

t.Commit()

bvalue_abvstd = 0.0

t = Transaction(doc, 'Place Hangers')
t.Start()
for dim in hanger.GetDimensions():
    if abs(delta_x) > abs(delta_y):
        Dimensions.append(dim.Name)
        if dim.Name == "Width":
            hanger.SetDimensionValue(dim, delta_y)
        if dim.Name == "Depth":
            hanger.SetDimensionValue(dim, 0)
        if dim.Name == "Bearer Extn":
            hanger.SetDimensionValue(dim, 0.33333333)
        hanger.get_Parameter(BuiltInParameter.FABRICATION_OFFSET_PARAM).Set(combined_bounding_box.Min.Z)
        translation = combined_bounding_box_Center - GetCenterPoint(hanger.Id)
        ElementTransformUtils.MoveElement(doc, hanger.Id, translation)
        hanger.get_Parameter(BuiltInParameter.FABRICATION_OFFSET_PARAM).Set(bottom_point)
    else:
        for dim in hanger.GetDimensions():
            if abs(delta_y) > abs(delta_x):
                Dimensions.append(dim.Name)
                if dim.Name == "Width":
                    hanger.SetDimensionValue(dim, delta_x)
                if dim.Name == "Depth":
                    hanger.SetDimensionValue(dim, 0)
                if dim.Name == "Bearer Extn":
                    hanger.SetDimensionValue(dim, 0.33333333)
                hanger.get_Parameter(BuiltInParameter.FABRICATION_OFFSET_PARAM).Set(combined_bounding_box.Min.Z)
                translation = combined_bounding_box_Center - GetCenterPoint(hanger.Id)
                ElementTransformUtils.MoveElement(doc, hanger.Id, translation)
                hanger.get_Parameter(BuiltInParameter.FABRICATION_OFFSET_PARAM).Set(bottom_point)
t.Commit()

t = Transaction(doc, "Round Trapeze Width")
t.Start()

if hanger:
    hanger.GetHostedInfo().DisconnectFromHost()
    for dim in hanger.GetDimensions():
        Dimensions.append(dim.Name)
        if dim.Name == "Width":
            width_value = hanger.GetDimensionValue(dim)
        if dim.Name == "Bearer Extn":
            bearer_value = hanger.GetDimensionValue(dim)
            in_bvalue = (bearer_value * 12)
            if in_bvalue > 4.0:
                bvalue_abvstd = in_bvalue - 4.0
            in_wvalue = (width_value * 12)
            rnd_value = myround((in_bvalue + in_wvalue + bvalue_abvstd), 2)
            abv_value = rnd_value - in_wvalue
            hlf_diff = (abv_value - 4.0) / 2
            new_value = (abv_value - hlf_diff) / 12
            hanger.SetDimensionValue(dim, new_value)
t.Commit()

