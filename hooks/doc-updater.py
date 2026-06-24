# -*- coding: utf-8 -*-
# doc-updater.py

import os
import clr

clr.AddReference('System')

from Autodesk.Revit.DB import (
    BuiltInCategory,
    ElementId,
    WorksharingUtils,
    FabricationPart,
    FabricationConfiguration
)

from Parameters.Get_Set_Params import (
    set_parameter_by_name,
    get_parameter_value_by_name_AsString,
    get_parameter_value_by_name_AsInteger,
    get_parameter_value_by_name_AsValueString,
    get_parameter_value_by_name_AsDouble
)

# ── Guard: bail out immediately if hook is disabled ───────────────────────────
FLAG_FILE = r'C:\temp\fabrication_hook_enabled.txt'

def hook_is_enabled():
    try:
        if os.path.exists(FLAG_FILE):
            with open(FLAG_FILE, 'r') as f:
                state = f.read().strip().lower()
                return state == 'true'
    except Exception:
        pass
    return False

if not hook_is_enabled():
    pass
else:
    sender = __eventsender__
    args = __eventargs__
    doc = args.GetDocument()
    Config = FabricationConfiguration.GetFabricationConfiguration(doc)

    modified_el_ids = list(args.GetModifiedElementIds())
    deleted_el_ids  = list(args.GetDeletedElementIds())
    new_el_ids      = list(args.GetAddedElementIds())

    combined_el_ids = [
        doc.GetElement(e_id)
        for e_id in modified_el_ids + new_el_ids
        if doc.GetElement(e_id) is not None
    ]

    allowed_cats = [
        ElementId(BuiltInCategory.OST_FabricationHangers),
        ElementId(BuiltInCategory.OST_FabricationDuctwork),
        ElementId(BuiltInCategory.OST_FabricationPipework),
        ElementId(BuiltInCategory.OST_FlexDuctCurves)
    ]

    def get_fabrication_connector_info(fabPart):
        connector_info = {}
        if isinstance(fabPart, FabricationPart):
            connectors = fabPart.ConnectorManager.Connectors
            for connector in connectors:
                try:
                    fab_connector_info = connector.GetFabricationConnectorInfo()
                    fabrication_conn_id = fab_connector_info.BodyConnectorId
                    connector_name = Config.GetFabricationConnectorName(fabrication_conn_id)
                    connector_info[connector.Id] = connector_name
                except Exception:
                    pass
        return connector_info

    for x in combined_el_ids:
        if x is None:
            continue

        if not hasattr(x, 'Category') or not x.Category or not hasattr(x.Category, 'Id'):
            continue

        if x.Category.Id not in allowed_cats:
            continue

        # ── FLEX DUCT LOGIC ────────────────────────────────────────────────────
        if x.Category.Id == ElementId(BuiltInCategory.OST_FlexDuctCurves):
            try:
                set_parameter_by_name(x, 'FP_Centerline Length', get_parameter_value_by_name_AsDouble(x, 'Length'))
            except:
                pass

            try:
                set_parameter_by_name(x, 'FP_Product Entry', get_parameter_value_by_name_AsString(x, 'Overall Size'))
            except:
                pass

            continue

        # ── FAB PARTS ONLY BELOW ───────────────────────────────────────────────
        if not isinstance(x, FabricationPart):
            continue

        try:
            part_material = get_parameter_value_by_name_AsValueString(x, 'Part Material')
            if part_material:
                set_parameter_by_name(x, 'FP_Part Material', part_material)
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
                if trimsize and trimangle:
                    set_parameter_by_name(x, 'FP_Product Entry', trimsize + ' x ' + trimangle)
        except:
            pass

        try:
            TOPE = 0.0
            BOTE = 0.0
            ItmDims = x.GetDimensions()
            for dta in ItmDims:
                if dta.Name == 'Top Extension':
                    TOPE = x.GetDimensionValue(dta) or 0.0
                if dta.Name == 'Bottom Extension':
                    BOTE = x.GetDimensionValue(dta) or 0.0
            set_parameter_by_name(x, 'FP_Extension Top', TOPE)
            set_parameter_by_name(x, 'FP_Extension Bottom', BOTE)
        except:
            pass

        # ── HANGER LOGIC ───────────────────────────────────────────────────────
        try:
            if x.ItemCustomId == 838:
                try:
                    rod_attached = 'Yes' if x.GetRodInfo().IsAttachedToStructure else 'No'
                    set_parameter_by_name(x, 'FP_Rod Attached', rod_attached)
                except:
                    pass

                try:
                    for n in x.GetPartAncillaryUsage():
                        if n.AncillaryWidthOrDiameter > 0:
                            set_parameter_by_name(x, 'FP_Rod Size', n.AncillaryWidthOrDiameter)
                except:
                    pass

                try:
                    if x.LookupParameter('Product Entry'):
                        set_parameter_by_name(x, 'FP_Hanger Diameter', get_parameter_value_by_name_AsString(x, 'Product Entry'))
                    else:
                        set_parameter_by_name(x, 'FP_Hanger Diameter', get_parameter_value_by_name_AsString(x, 'Size'))
                except:
                    pass

                RLA = 0.0
                RLB = 0.0
                TrapWidth = 0.0
                TrapExtn = 0.0

                if x.GetRodInfo().RodCount < 2:
                    hosted_info = x.GetHostedInfo().HostId
                    try:
                        host_el = doc.GetElement(hosted_info)
                        HostSize = get_parameter_value_by_name_AsString(host_el, 'Size').strip('"') if host_el else ''
                        HangerSize = get_parameter_value_by_name_AsString(x, 'Product Entry')
                        set_parameter_by_name(x, 'FP_Product Entry', HangerSize)

                        if HostSize == HangerSize and HostSize:
                            set_parameter_by_name(x, 'FP_Hanger Shield', 'No')
                            set_parameter_by_name(x, 'FP_Hanger Host Diameter', HostSize)
                        else:
                            set_parameter_by_name(x, 'FP_Hanger Shield', 'Yes')
                            if HostSize:
                                set_parameter_by_name(x, 'FP_Hanger Host Diameter', HostSize)
                    except:
                        pass

                    try:
                        ItmDims = x.GetDimensions()
                        for dta in ItmDims:
                            if dta.Name == 'Length A':
                                RLA = x.GetDimensionValue(dta) or 0.0
                            if dta.Name == 'Length B':
                                RLB = x.GetDimensionValue(dta) or 0.0

                        set_parameter_by_name(x, 'FP_Rod Length A', RLA)
                        set_parameter_by_name(x, 'FP_Rod Length B', RLB)
                    except:
                        pass

                else:
                    try:
                        ItmDims = x.GetDimensions()
                        for dta in ItmDims:
                            if dta.Name == 'Length A':
                                RLA = x.GetDimensionValue(dta) or 0.0
                            if dta.Name == 'Length B':
                                RLB = x.GetDimensionValue(dta) or 0.0
                            if dta.Name == 'Width':
                                TrapWidth = x.GetDimensionValue(dta) or 0.0
                            if dta.Name == 'Bearer Extn':
                                TrapExtn = x.GetDimensionValue(dta) or 0.0

                        BearerLength = TrapWidth + TrapExtn + TrapExtn
                        set_parameter_by_name(x, 'FP_Bearer Length', BearerLength)
                        set_parameter_by_name(x, 'FP_Rod Length', RLA)
                        set_parameter_by_name(x, 'FP_Rod Length A', RLA)
                        set_parameter_by_name(x, 'FP_Rod Length B', RLB)
                    except:
                        pass
        except:
            pass

        # ── CONNECTOR LOGIC ────────────────────────────────────────────────────
        try:
            connector_info = get_fabrication_connector_info(x)
            for connector_id, name in connector_info.items():
                param_name = "FP_Connector C" + str(connector_id + 1)
                set_parameter_by_name(x, param_name, name)
        except:
            pass

    # ── Optional: print checkout status for newly created elements ────────────
    for e_id in new_el_ids:
        element = doc.GetElement(e_id)
        if element:
            try:
                owner = WorksharingUtils.GetCheckoutStatus(doc, e_id)
                # print("Element ID: {}, Checkout Status: {}".format(e_id, owner))
            except:
                pass