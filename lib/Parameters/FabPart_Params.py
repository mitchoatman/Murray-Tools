def Sync_FP_Params_Entire_Model():
    import Autodesk
    from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, FabricationPart, FabricationConfiguration, WorksharingUtils, WorksetId
    from Autodesk.Revit.UI import TaskDialog
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

    # Check if the model is workshared
    if not doc.IsWorkshared:
        # Proceed without workset checkout logic if not workshared
        pass
    else:
        # Collect all elements to determine their worksets
        element_ids = set()  # Use a set to collect unique ElementIds
        collectors = [
            FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FabricationHangers).WhereElementIsNotElementType(),
            FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FabricationPipework).WhereElementIsNotElementType(),
            FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FabricationDuctwork).WhereElementIsNotElementType(),
            FilteredElementCollector(doc).OfClass(FabricationPart).WhereElementIsNotElementType(),
            FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FlexDuctCurves).WhereElementIsNotElementType()
        ]

        # Collect unique ElementIds from all collectors
        for collector in collectors:
            for element in collector.ToElements():
                element_ids.add(element.Id)

        # Convert ElementIds back to elements for workset checking
        all_elements = [doc.GetElement(eid) for eid in element_ids if doc.GetElement(eid)]

        # Get unique workset IDs for all elements
        workset_ids = set()
        for element in all_elements:
            try:
                workset_id = WorksharingUtils.GetWorksetId(doc, element.Id)
                if workset_id != WorksetId.InvalidWorksetId:
                    workset_ids.add(workset_id)
            except:
                pass

        # Attempt to check out worksets
        checkout_failed = False
        for workset_id in workset_ids:
            try:
                if not WorksharingUtils.GetCheckoutStatus(doc, workset_id) == DB.CheckoutStatus.OwnedByCurrentUser:
                    WorksharingUtils.CheckoutWorksets(doc, [workset_id])
            except Exception as e:
                checkout_failed = True
                break

        if checkout_failed:
            TaskDialog.Show(
                "Workset Checkout Failed",
                "Unable to check out one or more worksets. They may be owned by another user. Please try again later."
            )
            return

    # Existing collectors
    hanger_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FabricationHangers).WhereElementIsNotElementType().ToElements()
    pipe_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FabricationPipework).WhereElementIsNotElementType().ToElements()
    duct_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FabricationDuctwork).WhereElementIsNotElementType().ToElements()
    AllElements = FilteredElementCollector(doc).OfClass(FabricationPart).WhereElementIsNotElementType().ToElements()
    flex_duct_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FlexDuctCurves).WhereElementIsNotElementType().ToElements()

    # Start transaction
    t = Transaction(doc, "Update FP Parameters")
    t.Start()

    # Existing logic for ducts
    for duct in duct_collector:
        try:
            TOPE = 0
            BOTE = 0
            ItmDims = duct.GetDimensions()
            for dta in ItmDims:
                if dta.Name == 'Top Extension':
                    TOPE = duct.GetDimensionValue(dta)
                if dta.Name == 'Bottom Extension':
                    BOTE = duct.GetDimensionValue(dta)
            if TOPE and BOTE != 0:
                set_parameter_by_name(duct, 'FP_Extension Top', TOPE)
                set_parameter_by_name(duct, 'FP_Extension Bottom', BOTE)
        except:
            pass

    # Existing logic for hangers
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

    # Existing safely_set_parameter function
    def safely_set_parameter(action, elements):
        for element in elements:
            try:
                action(element)
            except Exception as e:
                pass

    # Existing actions
    actions = [
        lambda x: set_parameter_by_name(x, 'FP_Centerline Length', x.CenterlineLength) if x.ItemCustomId == 2041 else None,
        lambda x: set_parameter_by_name(x, 'FP_CID', x.ItemCustomId),
        lambda x: set_parameter_by_name(x, 'FP_Service Type', Config.GetServiceTypeName(x.ServiceType)),
        lambda x: set_parameter_by_name(x, 'FP_Service Name', get_parameter_value_by_name_AsString(x, 'Fabrication Service Name')),
        lambda x: set_parameter_by_name(x, 'FP_Service Abbreviation', get_parameter_value_by_name_AsString(x, 'Fabrication Service Abbreviation')),
        lambda x: set_parameter_by_name(x, 'FP_Rod Attached', 'Yes') if x.GetRodInfo().IsAttachedToStructure else set_parameter_by_name(x, 'FP_Rod Attached', 'No'),
        lambda x: [set_parameter_by_name(x, 'FP_Rod Size', n.AncillaryWidthOrDiameter) for n in x.GetPartAncillaryUsage() if n.AncillaryWidthOrDiameter > 0],
        lambda x: set_parameter_by_name(x, 'FP_Hanger Diameter', get_parameter_value_by_name_AsString(x, 'Product Entry')) if x.LookupParameter('Product Entry') else None,
        lambda x: set_parameter_by_name(x, 'FP_Product Entry', get_parameter_value_by_name_AsString(x, 'Product Entry')) if x.LookupParameter('Product Entry')
        else set_parameter_by_name(x, 'FP_Product Entry', get_parameter_value_by_name_AsString(x, 'Size')),
        lambda x: set_parameter_by_name(x, 'FP_Product Entry', (get_parameter_value_by_name_AsString(x, 'Size') or '') + ' x ' + (get_parameter_value_by_name_AsValueString(x, 'Angle') or '')) \
        if x.Alias and x.Alias.upper() == 'TRM' else set_parameter_by_name(x, 'FP_Product Entry', get_parameter_value_by_name_AsString(x, 'Size')),
        lambda x: set_parameter_by_name(x, 'FP_Centerline Length', x.CenterlineLength),
        lambda x: set_parameter_by_name(x, 'FP_Centerline Length', get_parameter_value_by_name_AsDouble(x, 'Length')),
        lambda x: set_parameter_by_name(x, 'FP_Product Entry', get_parameter_value_by_name_AsString(x, 'Overall Size')),
        lambda x: set_parameter_by_name(x, 'FP_Part Material', get_parameter_value_by_name_AsValueString(x, 'Part Material')) if get_parameter_value_by_name_AsValueString(x, 'Part Material') else None,
    ]

    # Apply actions
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

    # Existing connector info logic
    try:
        for x in AllElements:
            connector_info = get_fabrication_connector_info(x)
            for connector_id, name in connector_info.items():
                param_name = "FP_Connector C{}".format(connector_id + 1)
                set_parameter_by_name(x, param_name, name)
    except:
        pass

    t.Commit()

Sync_FP_Params_Entire_Model()