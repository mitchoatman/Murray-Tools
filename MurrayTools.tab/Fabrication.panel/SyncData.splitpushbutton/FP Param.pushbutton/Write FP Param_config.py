
import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, FabricationPart, FabricationConfiguration
from SharedParam.Add_Parameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString, get_parameter_value_by_name_AsInteger, get_parameter_value_by_name_AsValueString

Shared_Params()

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
fec = FilteredElementCollector
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float (RevitVersion)
Config = FabricationConfiguration.GetFabricationConfiguration(doc)

# Creating collector instance and collecting all the fabrication hangers from the model
hanger_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FabricationHangers) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

# Creating collector instance and collecting all the fabrication hangers from the model
pipe_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FabricationPipework) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

# Creating collector instance and collecting all the fabrication hangers from the model
duct_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FabricationDuctwork) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

# Create a FilteredElementCollector to get all FabricationPart elements
AllElements = FilteredElementCollector(doc).OfClass(FabricationPart) \
                   .WhereElementIsNotElementType() \
                   .ToElements()

t = Transaction(doc, "Update FP Parameters")
t.Start()

for hanger in hanger_collector:
    try:
        if (hanger.GetRodInfo().RodCount) < 2:
        
            # Get the host element's size
            hosted_info = hanger.GetHostedInfo().HostId
            HostSize = get_parameter_value_by_name_AsString(doc.GetElement(hosted_info), 'Size')
            # Get the hanger's size
            HangerSize = get_parameter_value_by_name_AsString(hanger, 'Product Entry')
            # Set the 'FP_Hanger Shield' parameter based on whether the host and hanger sizes match
            if HostSize == HangerSize:
                set_parameter_by_name(hanger, 'FP_Hanger Shield', 'No')
            else:
                set_parameter_by_name(hanger, 'FP_Hanger Shield', 'Yes')
                set_parameter_by_name(hanger, 'Comments', HostSize)  
                set_parameter_by_name(hanger, 'FP_Hanger Host Diameter', HostSize)                

            ItmDims = hanger.GetDimensions()
            for dta in ItmDims:
                if dta.Name == 'Rod Extn Below':
                    RB = hanger.GetDimensionValue(dta)
                if dta.Name == 'Rod Length':
                    RL = hanger.GetDimensionValue(dta)
            TRL = RL + RB
            set_parameter_by_name(hanger, 'FP_Rod Length', TRL)
    except:
        pass

    try:
        if (hanger.GetRodInfo().RodCount) > 1:
            ItmDims = hanger.GetDimensions()
            for dta in ItmDims:
                if dta.Name == 'Length A':
                    RLA = hanger.GetDimensionValue(dta)
                if dta.Name == 'Length B':
                    RLB = hanger.GetDimensionValue(dta)
                if dta.Name == 'Width':
                    TrapWidth = hanger.GetDimensionValue(dta)
                if dta.Name == 'Bearer Extn':
                    TrapExtn = hanger.GetDimensionValue(dta)
                if dta.Name == 'Right Rod Offset':
                    TrapRRod = hanger.GetDimensionValue(dta)
                if dta.Name == 'Left Rod Offset':
                    TrapLRod = hanger.GetDimensionValue(dta)
            BearerLength = TrapWidth + TrapExtn + TrapExtn
            set_parameter_by_name(hanger, 'FP_Bearer Length', BearerLength)
            set_parameter_by_name(hanger, 'FP_Rod Length', RLA)
            set_parameter_by_name(hanger, 'FP_Rod Length A', RLA)
            set_parameter_by_name(hanger, 'FP_Rod Length B', RLB)
    except:
        pass

def safely_set_parameter(action, elements):
    for element in elements:
        try:
            action(element)
        except Exception as e:
            # Log the exception if needed
            pass

# Define the actions as lambda functions
actions = [
    lambda x: set_parameter_by_name(x, 'FP_Centerline Length', x.CenterlineLength) if x.ItemCustomId == 2041 else None,
    lambda x: set_parameter_by_name(x, 'FP_Centerline Length', x.CenterlineLength),
    lambda x: set_parameter_by_name(x, 'FP_CID', x.ItemCustomId),
    lambda x: set_parameter_by_name(x, 'FP_Service Type', Config.GetServiceTypeName(x.ServiceType)),
    lambda x: set_parameter_by_name(x, 'FP_Service Name', get_parameter_value_by_name_AsString(x, 'Fabrication Service Name')),
    lambda x: set_parameter_by_name(x, 'FP_Service Abbreviation', get_parameter_value_by_name_AsString(x, 'Fabrication Service Abbreviation')),
    lambda x: set_parameter_by_name(x, 'FP_Rod Attached', 'Yes') if x.GetRodInfo().IsAttachedToStructure else set_parameter_by_name(x, 'FP_Rod Attached', 'No'),
    lambda x: [set_parameter_by_name(x, 'FP_Rod Size', n.AncillaryWidthOrDiameter) for n in x.GetPartAncillaryUsage() if n.AncillaryWidthOrDiameter > 0],
    lambda x: set_parameter_by_name(x, 'FP_Hanger Diameter', get_parameter_value_by_name_AsString(x, 'Product Entry')) if x.LookupParameter('Product Entry') else None,

    lambda x: set_parameter_by_name(x, 'FP_Product Entry', get_parameter_value_by_name_AsString(x, 'Product Entry')) if x.LookupParameter('Product Entry') \
    else set_parameter_by_name(x, 'FP_Product Entry', get_parameter_value_by_name_AsString(x, 'Size')),

    lambda x: set_parameter_by_name(x, 'FP_Product Entry', get_parameter_value_by_name_AsString(x, 'Product Entry')) if x.LookupParameter('Product Entry') \
    else set_parameter_by_name(x, 'FP_Product Entry', get_parameter_value_by_name_AsString(x, 'Size')) if x.Alias != 'TRM' \
    else set_parameter_by_name(x, 'FP_Product Entry', (get_parameter_value_by_name_AsString(x, 'Size') or '') + ' x ' + (get_parameter_value_by_name_AsValueString(x, 'Angle') or ''))


]

# Apply the actions to the respective element collections
safely_set_parameter(actions[0], AllElements)
safely_set_parameter(actions[1], pipe_collector)
safely_set_parameter(actions[2], duct_collector)
safely_set_parameter(actions[3], AllElements)
safely_set_parameter(actions[4], AllElements)
safely_set_parameter(actions[5], AllElements)
safely_set_parameter(actions[6], hanger_collector)
safely_set_parameter(actions[7], hanger_collector)
safely_set_parameter(actions[8], hanger_collector)
safely_set_parameter(actions[9], AllElements)
safely_set_parameter(actions[10], pipe_collector)

t.Commit()
