# -*- coding: utf-8 -*-
# doc-changed.event.py

import os
import clr
clr.AddReference('System')
from Autodesk.Revit.DB import (BuiltInCategory, ElementId, WorksharingUtils,
                                FabricationPart, FabricationConfiguration)
from Parameters.Get_Set_Params import (set_parameter_by_name,
                                        get_parameter_value_by_name_AsString,
                                        get_parameter_value_by_name_AsInteger,
                                        get_parameter_value_by_name_AsValueString)

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
    return False  # Default OFF if file is missing

if not hook_is_enabled():
    # Stop here — don't process anything
    import sys
    sys.exit(0)

# ── Hook is ON — proceed with full logic ──────────────────────────────────────
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
    ElementId(BuiltInCategory.OST_FabricationPipework)
]

def get_fabrication_connector_info(fabPart):
    connector_info = {}
    if isinstance(fabPart, FabricationPart):
        connectors = fabPart.ConnectorManager.Connectors
        for connector in connectors:
            try:
                fab_connector_info  = connector.GetFabricationConnectorInfo()
                fabrication_conn_id = fab_connector_info.BodyConnectorId
                connector_name      = Config.GetFabricationConnectorName(fabrication_conn_id)
                connector_info[connector.Id] = connector_name
            except Exception:
                print("Failed to get connector info for element {}".format(fabPart.Id))
    return connector_info

for x in combined_el_ids:
    if not isinstance(x, FabricationPart):
        continue
    if not hasattr(x, 'Category') or not x.Category or not hasattr(x.Category, 'Id'):
        continue
    if x.Category.Id in allowed_cats:
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

        if x.Alias == 'TRM':
            trimsize  = get_parameter_value_by_name_AsString(x, 'Size')
            trimangle = get_parameter_value_by_name_AsValueString(x, 'Angle')
            set_parameter_by_name(x, 'FP_Product Entry', trimsize + ' x ' + trimangle)

        if x.ItemCustomId == 838:
            rod_attached = 'Yes' if x.GetRodInfo().IsAttachedToStructure else 'No'
            set_parameter_by_name(x, 'FP_Rod Attached', rod_attached)

            for n in x.GetPartAncillaryUsage():
                if n.AncillaryWidthOrDiameter > 0:
                    set_parameter_by_name(x, 'FP_Rod Size', n.AncillaryWidthOrDiameter)

            if x.GetRodInfo().RodCount < 2:
                hosted_info = x.GetHostedInfo().HostId
                try:
                    HostSize = get_parameter_value_by_name_AsString(doc.GetElement(hosted_info), 'Size').strip('"')
                    HangerSize = get_parameter_value_by_name_AsString(x, 'Product Entry')
                    set_parameter_by_name(x, 'FP_Product Entry', HangerSize)
                    if HostSize == HangerSize:
                        set_parameter_by_name(x, 'FP_Hanger Shield', 'No')
                        set_parameter_by_name(x, 'FP_Hanger Host Diameter', HostSize)
                    else:
                        set_parameter_by_name(x, 'FP_Hanger Shield', 'Yes')
                except:
                    pass

                ItmDims = x.GetDimensions()
                for dta in ItmDims:
                    if dta.Name == 'Length A':
                        RLA = x.GetDimensionValue(dta) or 0.0
                    if dta.Name == 'Length B':
                        RLB = x.GetDimensionValue(dta) or 0.0
                set_parameter_by_name(x, 'FP_Rod Length', RLA)
                set_parameter_by_name(x, 'FP_Rod Length A', RLA)
                set_parameter_by_name(x, 'FP_Rod Length B', RLB)

            else:
                ItmDims = x.GetDimensions()
                for dta in ItmDims:
                    if dta.Name == 'Length A':
                        RLA = x.GetDimensionValue(dta) or 0.0
                    if dta.Name == 'Length B':
                        RLB = x.GetDimensionValue(dta) or 0.0
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


        connector_info = get_fabrication_connector_info(x)
        for connector_id, name in connector_info.items():
            param_name = "FP_Connector C" + str(connector_id + 1)
            set_parameter_by_name(x, param_name, name)

for e_id in new_el_ids:
    element = doc.GetElement(e_id)
    if element:
        owner = WorksharingUtils.GetCheckoutStatus(doc, e_id)
        print("Element ID: {}, Checkout Status: {}".format(e_id, owner))