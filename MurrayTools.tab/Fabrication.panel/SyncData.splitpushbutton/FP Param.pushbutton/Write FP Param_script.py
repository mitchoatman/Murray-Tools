
import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, FabricationPart, FabricationConfiguration, BuiltInParameter
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

# Function to get fabrication connector names and track their IDs
def get_fabrication_connector_info(fabPart):
    # Create a dictionary to store connector names with their IDs
    connector_info = {}
    # Check if the element is a fabrication part
    if isinstance(fabPart, FabricationPart):
        connectors = fabPart.ConnectorManager.Connectors
        # Loop through each connector
        for connector in connectors:
            try:
                # Get FabricationConnectorInfo from the connector
                fab_connector_info = connector.GetFabricationConnectorInfo()
                # Extract the BodyConnectorId
                fabrication_connector_id = fab_connector_info.BodyConnectorId
                # Use GetFabricationConnectorName with the correct fabrication configuration and connector ID
                connector_name = FabricationConfiguration.GetFabricationConfiguration(doc).GetFabricationConnectorName(fabrication_connector_id)
                # Store the connector ID and name in the dictionary
                connector_info[connector.Id] = connector_name  # Directly use connector.Id
            except:
                pass
    return connector_info


# --------------SELECTED ELEMENTS-------------------

selection = [doc.GetElement(id) for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]

if selection:
    t = Transaction(doc, "Update FP Parameters")
    t.Start()
    for x in selection:
        isfabpart = x.LookupParameter("Fabrication Service")
        if isfabpart:
            set_parameter_by_name(x, 'FP_Centerline Length', x.CenterlineLength) if x.ItemCustomId == 2041 else None
            set_parameter_by_name(x, 'FP_Centerline Length', x.CenterlineLength) if x.Category.Name == 'MEP Fabrication Ductwork' else None
            set_parameter_by_name(x, 'FP_CID', x.ItemCustomId)
            set_parameter_by_name(x, 'FP_Service Type', Config.GetServiceTypeName(x.ServiceType))
            set_parameter_by_name(x, 'FP_Service Name', get_parameter_value_by_name_AsString(x, 'Fabrication Service Name'))
            # Set the parameter only if it has a value
            service_abbreviation = get_parameter_value_by_name_AsString(x, 'Fabrication Service Abbreviation')
            if service_abbreviation:
                set_parameter_by_name(x, 'FP_Service Abbreviation', service_abbreviation)
            try:
                [set_parameter_by_name(x, 'FP_Product Entry', get_parameter_value_by_name_AsString(x, 'Product Entry')) if x.LookupParameter('Product Entry') else set_parameter_by_name(x, 'FP_Product Entry', get_parameter_value_by_name_AsString(x, 'Size'))]
                if x.Alias == 'TRM':
                    trimsize = get_parameter_value_by_name_AsString(x, 'Size')
                    trimangle = get_parameter_value_by_name_AsValueString(x, 'Angle')
                    set_parameter_by_name(x, 'FP_Product Entry', trimsize + ' x ' + trimangle)
            except:
                pass
            try:
                #------#
                ItmDims = x.GetDimensions()
                for dta in ItmDims:
                    if dta.Name == 'Top Extension':
                        TOPE = x.GetDimensionValue(dta)
                    if dta.Name == 'Bottom Extension':
                        BOTE = x.GetDimensionValue(dta)
                set_parameter_by_name(x, 'FP_Extension Top', TOPE)
                set_parameter_by_name(x, 'FP_Extension Bottom', BOTE)
            except:
                pass
            try:
                if x.ItemCustomId == 838:
                    set_parameter_by_name(x, 'FP_Rod Attached', 'Yes') if x.GetRodInfo().IsAttachedToStructure else set_parameter_by_name(x, 'FP_Rod Attached', 'No')
                    [set_parameter_by_name(x, 'FP_Rod Size', n.AncillaryWidthOrDiameter) for n in x.GetPartAncillaryUsage() if n.AncillaryWidthOrDiameter > 0]
                    ProductEntry = x.LookupParameter('Product Entry')
                    if ProductEntry:
                        if (x.GetRodInfo().RodCount) < 2:
                            # Get the host element's size
                            hosted_info = x.GetHostedInfo().HostId
                            try:
                                HostSize = get_parameter_value_by_name_AsString(doc.GetElement(hosted_info), 'Size').strip('"')
                                # Get the hanger's size
                                HangerSize = get_parameter_value_by_name_AsString(x, 'Product Entry')
                                set_parameter_by_name(x, 'FP_Product Entry', HangerSize)
                                # Set the 'FP_Hanger Shield' parameter based on whether the host and hanger sizes match
                                if HostSize == HangerSize:
                                    set_parameter_by_name(x, 'FP_Hanger Shield', 'No')
                                else:
                                    set_parameter_by_name(x, 'FP_Hanger Shield', 'Yes')
                                    set_parameter_by_name(x, 'Comments', HostSize)
                                    set_parameter_by_name(x, 'FP_Hanger Host Diameter', HostSize)
                            except:
                                pass

                            ItmDims = x.GetDimensions()
                            for dta in ItmDims:
                                if dta.Name == 'Length A':
                                    RLA = x.GetDimensionValue(dta)
                                if dta.Name == 'Length B':
                                    RLB = x.GetDimensionValue(dta)
                            set_parameter_by_name(x, 'FP_Rod Length', RLA)
                            set_parameter_by_name(x, 'FP_Rod Length A', RLA)
                            set_parameter_by_name(x, 'FP_Rod Length B', RLB)

                    else:
                        ItmDims = x.GetDimensions()
                        for dta in ItmDims:
                            if dta.Name == 'Length A':
                                RLA = x.GetDimensionValue(dta)
                            if dta.Name == 'Length B':
                                RLB = x.GetDimensionValue(dta)
                            if dta.Name == 'Width':
                                TrapWidth = x.GetDimensionValue(dta)
                            if dta.Name == 'Bearer Extn':
                                TrapExtn = x.GetDimensionValue(dta)
                            if dta.Name == 'Right Rod Offset':
                                TrapRRod = x.GetDimensionValue(dta)
                            if dta.Name == 'Left Rod Offset':
                                TrapLRod = x.GetDimensionValue(dta)
                        BearerLength = TrapWidth + TrapExtn + TrapExtn
                        set_parameter_by_name(x, 'FP_Bearer Length', BearerLength)
                        set_parameter_by_name(x, 'FP_Rod Length', RLA)
                        set_parameter_by_name(x, 'FP_Rod Length A', RLA)
                        set_parameter_by_name(x, 'FP_Rod Length B', RLB)
            except:
                pass
            try:
                if isinstance(x, FabricationPart):
                    connector_info = get_fabrication_connector_info(x)
                    for connector_id, name in connector_info.items():
                        param_name = "FP_Connector C{}".format(connector_id + 1)  # Offset by 1
                        set_parameter_by_name(x, param_name, name)
            except:
                pass
    t.Commit()

# --------------SELECTED ELEMENTS-------------------
else:
    # --------------ACTIVE VIEW-------------------

    # Creating collector instance and collecting all the fabrication hangers from the model
    hanger_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationHangers) \
                       .WhereElementIsNotElementType() \
                       .ToElements()

    # Creating collector instance and collecting all the fabrication hangers from the model
    pipe_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_PipeAccessory) \
                       .WhereElementIsNotElementType() \
                       .ToElements()

    # Creating collector instance and collecting all the fabrication hangers from the model
    duct_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationDuctwork) \
                       .WhereElementIsNotElementType() \
                       .ToElements()

    # Create a FilteredElementCollector to get all FabricationPart elements
    AllElements = FilteredElementCollector(doc, curview.Id).OfClass(FabricationPart) \
                       .WhereElementIsNotElementType() \
                       .ToElements()

    t = Transaction(doc, "Update FP Parameters")
    t.Start()

    for duct in duct_collector:
        try:
            ItmDims = duct.GetDimensions()
            for dta in ItmDims:
                if dta.Name == 'Top Extension':
                    TOPE = duct.GetDimensionValue(dta)
                if dta.Name == 'Bottom Extension':
                    BOTE = duct.GetDimensionValue(dta)
                set_parameter_by_name(duct, 'FP_Extension Top', TOPE)
                set_parameter_by_name(duct, 'FP_Extension Bottom', BOTE)
        except:
            pass

    for hanger in hanger_collector:
        if (hanger.GetRodInfo().RodCount) < 2:
            # Get the host element's size
            hosted_info = hanger.GetHostedInfo().HostId
            try:
                HostSize = get_parameter_value_by_name_AsString(doc.GetElement(hosted_info), 'Size').strip('"')
                # Get the hanger's size
                HangerSize = get_parameter_value_by_name_AsString(hanger, 'Product Entry')
                set_parameter_by_name(hanger, 'FP_Product Entry', HangerSize)
                set_parameter_by_name(hanger, 'Comments', HostSize)
                set_parameter_by_name(hanger, 'FP_Hanger Host Diameter', HostSize)
                # Set the 'FP_Hanger Shield' parameter based on whether the host and hanger sizes match
                if HostSize == HangerSize:
                    set_parameter_by_name(hanger, 'FP_Hanger Shield', 'No')
                else:
                    set_parameter_by_name(hanger, 'FP_Hanger Shield', 'Yes')
            except:
                pass
            ItmDims = hanger.GetDimensions()
            for dta in ItmDims:
                if dta.Name == 'Length A':
                    RLA = hanger.GetDimensionValue(dta)
            set_parameter_by_name(hanger, 'FP_Rod Length', RLA)
            set_parameter_by_name(hanger, 'FP_Rod Length A', RLA)

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
        else set_parameter_by_name(x, 'FP_Product Entry', (get_parameter_value_by_name_AsString(x, 'Size') or '') + ' x ' + (get_parameter_value_by_name_AsValueString(x, 'Angle') or '')),
        lambda x: set_parameter_by_name(x, 'FP_Centerline Length', x.CenterlineLength),
        ]

    # Apply the actions to the respective element collections
    safely_set_parameter(actions[0], AllElements)
    safely_set_parameter(actions[1], AllElements)
    safely_set_parameter(actions[2], AllElements)
    safely_set_parameter(actions[3], AllElements)
    safely_set_parameter(actions[4], AllElements)
    safely_set_parameter(actions[5], hanger_collector)
    safely_set_parameter(actions[6], hanger_collector)
    safely_set_parameter(actions[7], hanger_collector)
    safely_set_parameter(actions[8], AllElements)
    safely_set_parameter(actions[9], pipe_collector)
    safely_set_parameter(actions[10], duct_collector)

    try:
        for x in AllElements:
            connector_info = get_fabrication_connector_info(x)
            for connector_id, name in connector_info.items():
                param_name = "FP_Connector C{}".format(connector_id + 1)  # Offset by 1
                set_parameter_by_name(x, param_name, name)
    except:
        pass

    t.Commit()
    
    # --------------ACTIVE VIEW-------------------