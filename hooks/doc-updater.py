# -*- coding: utf-8 -*-
import os
import clr
clr.AddReference('System')
import System
from Autodesk.Revit.DB import (
    BuiltInCategory,
    ElementId,
    FabricationPart,
    FabricationConfiguration,
    StorageType
)
from Parameters.Get_Set_Params import (
    set_parameter_by_name,
    get_parameter_value_by_name_AsString,
    get_parameter_value_by_name_AsInteger,
    get_parameter_value_by_name_AsValueString,
    get_parameter_value_by_name_AsDouble
)

# ── Settings ─────────────────────────────────────────────────────────────────
FLAG_FILE = r'C:\temp\Ribbon_fabrication-hook-status.txt'
LOG_FILE = r'C:\temp\Ribbon_fabrication-hook.log'

# Re-entrancy flag. Hook scripts re-run from scratch on every event, so a module
# variable cannot remember anything between runs. AppDomain data survives.
BUSY_KEY = "MurrayFabHookBusy"

# Transactions from other tools that this hook must never react to.
# Add the exact transaction name of any other tool that should be ignored.
IGNORED_TRANSACTIONS = (
    "WebView2 Batch View Rename",
)

ALLOWED_CATS = [
    ElementId(BuiltInCategory.OST_FabricationHangers),
    ElementId(BuiltInCategory.OST_FabricationDuctwork),
    ElementId(BuiltInCategory.OST_FabricationPipework),
    ElementId(BuiltInCategory.OST_FlexDuctCurves)
]


# ── Helpers ──────────────────────────────────────────────────────────────────
def hook_is_enabled():
    try:
        if os.path.exists(FLAG_FILE):
            with open(FLAG_FILE, 'r') as f:
                return f.read().strip().lower() == 'true'
    except Exception:
        pass
    return False


def log(msg):
    """Errors go to a file. Nothing from this hook may surface as a dialog,
    and print() from an event hook is not safe."""
    try:
        with open(LOG_FILE, 'a') as f:
            f.write(str(msg) + "\n")
    except Exception:
        pass


def safe_set_parameter(element, param_name, new_value):
    """Safely set parameter only if value differs to reduce event triggers."""
    if element is None:
        return
    try:
        param = element.LookupParameter(param_name)
        if not param:
            return

        storage_type = param.StorageType
        if storage_type == StorageType.String:
            current = param.AsString() or ""
            new_value_str = str(new_value) if new_value is not None else ""
            if current != new_value_str:
                set_parameter_by_name(element, param_name, new_value)
        elif storage_type == StorageType.Double:
            current = param.AsDouble()
            new_val_float = float(new_value) if new_value is not None else 0.0
            if current is None or abs(current - new_val_float) > 1e-6:
                set_parameter_by_name(element, param_name, new_value)
        elif storage_type == StorageType.Integer:
            current = param.AsInteger()
            new_val_int = int(new_value) if new_value is not None else 0
            if current != new_val_int:
                set_parameter_by_name(element, param_name, new_value)
        else:
            current = param.AsValueString() or ""
            if str(current) != str(new_value):
                set_parameter_by_name(element, param_name, new_value)
    except Exception:
        try:
            set_parameter_by_name(element, param_name, new_value)
        except Exception:
            pass


def get_fabrication_connector_info(fabPart, Config):
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


def process_element(x, doc, Config):
    # ── FLEX DUCT LOGIC ──────────────────────────────────────────────────────
    if x.Category.Id == ElementId(BuiltInCategory.OST_FlexDuctCurves):
        try:
            safe_set_parameter(x, 'FP_Centerline Length', get_parameter_value_by_name_AsDouble(x, 'Length'))
        except Exception:
            pass
        try:
            safe_set_parameter(x, 'FP_Product Entry', get_parameter_value_by_name_AsString(x, 'Overall Size'))
        except Exception:
            pass
        return

    # ── FAB PARTS ONLY BELOW ─────────────────────────────────────────────────
    if not isinstance(x, FabricationPart):
        return

    try:
        part_material = get_parameter_value_by_name_AsValueString(x, 'Part Material')
        if part_material:
            safe_set_parameter(x, 'FP_Part Material', part_material)
    except Exception:
        pass

    try:
        if x.ItemCustomId == 2041 or x.Category.Name == 'MEP Fabrication Ductwork':
            safe_set_parameter(x, 'FP_Centerline Length', x.CenterlineLength)
    except Exception:
        pass

    try:
        safe_set_parameter(x, 'FP_CID', x.ItemCustomId)
    except Exception:
        pass

    try:
        safe_set_parameter(x, 'FP_Service Type', Config.GetServiceTypeName(x.ServiceType))
    except Exception:
        pass

    try:
        safe_set_parameter(x, 'FP_Service Name', get_parameter_value_by_name_AsString(x, 'Fabrication Service Name'))
    except Exception:
        pass

    try:
        service_abbreviation = get_parameter_value_by_name_AsString(x, 'Fabrication Service Abbreviation')
        if service_abbreviation:
            safe_set_parameter(x, 'FP_Service Abbreviation', service_abbreviation)
    except Exception:
        pass

    try:
        if x.LookupParameter('Product Entry'):
            safe_set_parameter(x, 'FP_Product Entry', get_parameter_value_by_name_AsString(x, 'Product Entry'))
        else:
            safe_set_parameter(x, 'FP_Product Entry', get_parameter_value_by_name_AsString(x, 'Size'))
    except Exception:
        pass

    try:
        if x.Alias == 'TRM':
            trimsize = get_parameter_value_by_name_AsString(x, 'Size')
            trimangle = get_parameter_value_by_name_AsValueString(x, 'Angle')
            if trimsize and trimangle:
                safe_set_parameter(x, 'FP_Product Entry', trimsize + ' x ' + trimangle)
    except Exception:
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
        safe_set_parameter(x, 'FP_Extension Top', TOPE)
        safe_set_parameter(x, 'FP_Extension Bottom', BOTE)
    except Exception:
        pass

    # ── HANGER LOGIC ─────────────────────────────────────────────────────────
    try:
        if x.ItemCustomId == 838:
            try:
                rod_attached = 'Yes' if x.GetRodInfo().IsAttachedToStructure else 'No'
                safe_set_parameter(x, 'FP_Rod Attached', rod_attached)
            except Exception:
                pass
            try:
                for n in x.GetPartAncillaryUsage():
                    if n.AncillaryWidthOrDiameter > 0:
                        safe_set_parameter(x, 'FP_Rod Size', n.AncillaryWidthOrDiameter)
            except Exception:
                pass
            try:
                if x.LookupParameter('Product Entry'):
                    safe_set_parameter(x, 'FP_Hanger Diameter', get_parameter_value_by_name_AsString(x, 'Product Entry'))
                else:
                    safe_set_parameter(x, 'FP_Hanger Diameter', get_parameter_value_by_name_AsString(x, 'Size'))
            except Exception:
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
                    safe_set_parameter(x, 'FP_Product Entry', HangerSize)
                    if HostSize == HangerSize and HostSize:
                        safe_set_parameter(x, 'FP_Hanger Shield', 'No')
                        safe_set_parameter(x, 'FP_Hanger Host Diameter', HostSize)
                    else:
                        safe_set_parameter(x, 'FP_Hanger Shield', 'Yes')
                        if HostSize:
                            safe_set_parameter(x, 'FP_Hanger Host Diameter', HostSize)
                except Exception:
                    pass
                try:
                    ItmDims = x.GetDimensions()
                    for dta in ItmDims:
                        if dta.Name == 'Length A':
                            RLA = x.GetDimensionValue(dta) or 0.0
                        if dta.Name == 'Length B':
                            RLB = x.GetDimensionValue(dta) or 0.0
                    safe_set_parameter(x, 'FP_Rod Length A', RLA)
                    safe_set_parameter(x, 'FP_Rod Length B', RLB)
                except Exception:
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
                    safe_set_parameter(x, 'FP_Bearer Length', BearerLength)
                    safe_set_parameter(x, 'FP_Rod Length', RLA)
                    safe_set_parameter(x, 'FP_Rod Length A', RLA)
                    safe_set_parameter(x, 'FP_Rod Length B', RLB)
                except Exception:
                    pass
    except Exception:
        pass

    # ── CONNECTOR LOGIC ──────────────────────────────────────────────────────
    try:
        connector_info = get_fabrication_connector_info(x, Config)
        for connector_id, name in connector_info.items():
            param_name = "FP_Connector C" + str(connector_id + 1)
            safe_set_parameter(x, param_name, name)
    except Exception:
        pass


# ── Hook entry point ─────────────────────────────────────────────────────────
def run_hook():
    if not hook_is_enabled():
        return

    args = __eventargs__
    doc = args.GetDocument()
    if doc is None:
        return

    # 1. Ignore transactions that belong to other tools (for example the view renamer).
    #    Skipped only when every transaction in this change is on the ignore list.
    try:
        txn_names = list(args.GetTransactionNames())
    except Exception:
        txn_names = []
    if txn_names and all(n in IGNORED_TRANSACTIONS for n in txn_names):
        return

    # 2. Re-entrancy guard that actually persists between hook runs.
    if System.AppDomain.CurrentDomain.GetData(BUSY_KEY):
        return

    # 3. Look at what changed BEFORE doing any work. Keep only fabrication elements.
    #    Anything else (views, sheets, renames, text, etc.) exits here untouched.
    changed_ids = list(args.GetModifiedElementIds()) + list(args.GetAddedElementIds())
    if not changed_ids:
        return

    fab_elements = []
    for e_id in changed_ids:
        el = doc.GetElement(e_id)
        if el is None:
            continue
        cat = el.Category
        if cat is None:
            continue
        if cat.Id in ALLOWED_CATS:
            fab_elements.append(el)

    if not fab_elements:
        return

    # 4. Only now touch the fabrication configuration and write parameters.
    System.AppDomain.CurrentDomain.SetData(BUSY_KEY, True)
    try:
        Config = None
        try:
            Config = FabricationConfiguration.GetFabricationConfiguration(doc)
        except Exception as ex:
            log("No fabrication configuration: " + str(ex))

        for x in fab_elements:
            try:
                process_element(x, doc, Config)
            except Exception as ex:
                log("Element error: " + str(ex))
    finally:
        System.AppDomain.CurrentDomain.SetData(BUSY_KEY, None)


# Nothing is allowed to escape this hook. An uncaught error here shows up as a
# pyRevit error dialog in the middle of whatever the user was doing.
try:
    run_hook()
except Exception as _ex:
    log("Hook error: " + str(_ex))