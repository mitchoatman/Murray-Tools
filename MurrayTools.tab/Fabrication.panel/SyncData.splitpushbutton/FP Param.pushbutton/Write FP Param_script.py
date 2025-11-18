import Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, FabricationPart, FabricationConfiguration, BuiltInParameter, WorksharingUtils
from System.Collections.Generic import List
from Parameters.Add_SharedParameters import Shared_Params
from Parameters.Get_Set_Params import set_parameter_by_name, get_parameter_value_by_name_AsString, get_parameter_value_by_name_AsInteger, get_parameter_value_by_name_AsValueString, get_parameter_value_by_name_AsDouble

Shared_Params()

DB = Autodesk.Revit.DB
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
curview = doc.ActiveView
fec = FilteredElementCollector
app = doc.Application
RevitVersion = app.VersionNumber
RevitINT = float(RevitVersion)
Config = FabricationConfiguration.GetFabricationConfiguration(doc)

# Function to get fabrication connector names and track their IDs
def get_fabrication_connector_info(fabPart):
    connector_info = {}
    if isinstance(fabPart, FabricationPart):
        connectors = fabPart.ConnectorManager.Connectors
        for connector in connectors:
            try:
                fab_connector_info = connector.GetFabricationConnectorInfo()
                fabrication_connector_id = fab_connector_info.BodyConnectorId
                connector_name = FabricationConfiguration.GetFabricationConfiguration(doc).GetFabricationConnectorName(fabrication_connector_id)
                connector_info[connector.Id] = connector_name
            except:
                pass
    return connector_info

# Function to check out elements in one operation
def checkout_elements(element_ids):
    if not doc.IsWorkshared or not element_ids:
        return
    try:
        # Convert to List[ElementId] for CheckoutElements
        id_list = List[DB.ElementId](element_ids)
        WorksharingUtils.CheckoutElements(doc, id_list)
    except:
        pass  # If checkout fails, proceed with editable elements only

# Collect all elements that will be modified and check for checkout
all_elements_to_modify = []
selection = [doc.GetElement(id) for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
if selection:
    all_elements_to_modify.extend(selection)
else:
    hanger_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationHangers) \
                       .WhereElementIsNotElementType() \
                       .ToElements()
    pipe_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationPipework) \
                       .WhereElementIsNotElementType() \
                       .ToElements()
    duct_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationDuctwork) \
                       .WhereElementIsNotElementType() \
                       .ToElements()
    AllElements = FilteredElementCollector(doc, curview.Id).OfClass(FabricationPart) \
                       .WhereElementIsNotElementType() \
                       .ToElements()
    flex_duct_collector = FilteredElementCollector(doc, curview.Id) \
                            .OfCategory(BuiltInCategory.OST_FlexDuctCurves) \
                            .WhereElementIsNotElementType() \
                            .ToElements()
    all_elements_to_modify.extend(hanger_collector)
    all_elements_to_modify.extend(pipe_collector)
    all_elements_to_modify.extend(duct_collector)
    all_elements_to_modify.extend(AllElements)
    all_elements_to_modify.extend(flex_duct_collector)

# Identify non-editable elements and attempt to check them out
non_editable_ids = []
if doc.IsWorkshared:
    for elem in all_elements_to_modify:
        try:
            tooltip_info = WorksharingUtils.GetWorksharingTooltipInfo(doc, elem.Id)
            if not tooltip_info.Editable:
                non_editable_ids.append(elem.Id)
        except:
            pass
if non_editable_ids:
    checkout_elements(non_editable_ids)

# --------------SELECTED ELEMENTS-------------------
if selection:
    t = Transaction(doc, "Update FP Parameters")
    t.Start()
    for x in selection:
        isfabpart = x.LookupParameter("Fabrication Service")
        if isfabpart:
            try:
                if x.ItemCustomId != 838:
                    set_parameter_by_name(x, 'FP_Part Material', get_parameter_value_by_name_AsValueString(x, 'Part Material'))
            except:
                pass

            try:
                if x.ItemCustomId == 2041 or x.Category.Name == 'MEP Fabrication Ductwork':
                    set_parameter_by_name(x, 'FP_Centerline Length', x.CenterlineLength)
            except:
                pass

            try:
                set_parameter_by_name(x, 'FP_CID', x.ItemCustomId)
            except:
                pass

            try:
                set_parameter_by_name(x, 'FP_Service Type', Config.GetServiceTypeName(x.ServiceType))
            except:
                pass

            try:
                set_parameter_by_name(x, 'FP_Service Name', get_parameter_value_by_name_AsString(x, 'Fabrication Service Name'))
            except:
                pass

            try:
                service_abbreviation = get_parameter_value_by_name_AsString(x, 'Fabrication Service Abbreviation')
                if service_abbreviation:
                    set_parameter_by_name(x, 'FP_Service Abbreviation', service_abbreviation)
            except:
                pass

            try:
                if x.LookupParameter('Product Entry'):
                    set_parameter_by_name(x, 'FP_Product Entry', get_parameter_value_by_name_AsString(x, 'Product Entry'))
                else:
                    set_parameter_by_name(x, 'FP_Product Entry', get_parameter_value_by_name_AsString(x, 'Size'))
            except:
                pass

            try:
                if x.Alias == 'TRM':
                    trimsize = get_parameter_value_by_name_AsString(x, 'Size')
                    trimangle = get_parameter_value_by_name_AsValueString(x, 'Angle')
                    set_parameter_by_name(x, 'FP_Product Entry', trimsize + ' x ' + trimangle)
            except:
                pass
            try:
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
                            hosted_info = x.GetHostedInfo().HostId
                            try:
                                HostSize = get_parameter_value_by_name_AsString(doc.GetElement(hosted_info), 'Size').strip('"')
                                HangerSize = get_parameter_value_by_name_AsString(x, 'Product Entry')
                                set_parameter_by_name(x, 'FP_Product Entry', HangerSize)
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
                        param_name = "FP_Connector C{}".format(connector_id + 1)
                        set_parameter_by_name(x, param_name, name)
            except:
                pass
    t.Commit()

# --------------ACTIVE VIEW-------------------
else:
    hanger_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationHangers) \
                       .WhereElementIsNotElementType() \
                       .ToElements()
    pipe_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationPipework) \
                       .WhereElementIsNotElementType() \
                       .ToElements()
    duct_collector = FilteredElementCollector(doc, curview.Id).OfCategory(BuiltInCategory.OST_FabricationDuctwork) \
                       .WhereElementIsNotElementType() \
                       .ToElements()
    AllElements = FilteredElementCollector(doc, curview.Id).OfClass(FabricationPart) \
                       .WhereElementIsNotElementType() \
                       .ToElements()
    flex_duct_collector = FilteredElementCollector(doc, curview.Id) \
                            .OfCategory(BuiltInCategory.OST_FlexDuctCurves) \
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
            hosted_info = hanger.GetHostedInfo().HostId
            try:
                HostSize = get_parameter_value_by_name_AsString(doc.GetElement(hosted_info), 'Size').strip('"')
                HangerSize = get_parameter_value_by_name_AsString(hanger, 'Product Entry')
                set_parameter_by_name(hanger, 'FP_Product Entry', HangerSize)
                set_parameter_by_name(hanger, 'Comments', HostSize)
                set_parameter_by_name(hanger, 'FP_Hanger Host Diameter', HostSize)
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
                pass

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
        lambda x: set_parameter_by_name(x, 'FP_Product Entry', (get_parameter_value_by_name_AsString(x, 'Size') or '') + ' x ' + (get_parameter_value_by_name_AsValueString(x, 'Angle') \
        or '')) if x.Alias and x.Alias.upper() == 'TRM' else get_parameter_value_by_name_AsString(x, 'Size'),
        lambda x: set_parameter_by_name(x, 'FP_Centerline Length', x.CenterlineLength),
        lambda x: set_parameter_by_name(x, 'FP_Centerline Length', get_parameter_value_by_name_AsDouble(x, 'Length')),
        lambda x: set_parameter_by_name(x, 'FP_Product Entry', get_parameter_value_by_name_AsString(x, 'Overall Size')),
        lambda x: set_parameter_by_name(x, 'FP_Part Material', get_parameter_value_by_name_AsValueString(x, 'Part Material')) if get_parameter_value_by_name_AsValueString(x, 'Part Material') else None,
    ]

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
    safely_set_parameter(actions[11], flex_duct_collector)
    safely_set_parameter(actions[12], flex_duct_collector)
    safely_set_parameter(actions[13], pipe_collector)
    safely_set_parameter(actions[13], duct_collector)

    try:
        for x in AllElements:
            connector_info = get_fabrication_connector_info(x)
            for connector_id, name in connector_info.items():
                param_name = "FP_Connector C{}".format(connector_id + 1)
                set_parameter_by_name(x, param_name, name)
    except:
        pass

    t.Commit()