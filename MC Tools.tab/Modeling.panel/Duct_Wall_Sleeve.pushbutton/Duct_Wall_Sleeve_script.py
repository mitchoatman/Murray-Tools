# -*- coding: utf-8 -*-
import clr
clr.AddReference('System')
import System

from Autodesk.Revit import DB
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Events import TaskDialogShowingEventArgs
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Family,
    LocationCurve,
    Transaction
)
from Autodesk.Revit.Exceptions import OperationCanceledException

from Parameters.Get_Set_Params import (
    set_parameter_by_name,
    get_parameter_value_by_name_AsString
)
from Parameters.Add_SharedParameters import Shared_Params

import os
from math import atan2

Shared_Params()

# --------------------------------------------------
# Basic environment
# --------------------------------------------------
path, filename = os.path.split(__file__)
round_family_path = os.path.join(path, 'Round Wall Sleeve.rfa')
rect_family_path = os.path.join(path, 'Rectangular Wall Sleeve.rfa')

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
uiapp = __revit__
active_view = doc.ActiveView

try:
    level = active_view.GenLevel
except:
    level = None

# --------------------------------------------------
# Helpers
# --------------------------------------------------

def safe_get_double_param_by_name(element, param_name, default=0.0):
    try:
        if element is None or not param_name:
            return default

        param = element.LookupParameter(param_name)
        if not param:
            return default

        if param.StorageType == DB.StorageType.Double:
            return param.AsDouble()

        if param.StorageType == DB.StorageType.Integer:
            return float(param.AsInteger())

        if param.StorageType == DB.StorageType.String:
            val = param.AsString()
            if val:
                return float(val)

    except:
        pass

    return default


def get_insulation_thickness(duct):
    """
    Returns insulation thickness in feet.
    Tries common parameter names and falls back to 0.
    """
    possible_names = [
        'Insulation Thickness',
        'Insulation',
        'Ins Thickness'
    ]

    for pname in possible_names:
        val = safe_get_double_param_by_name(duct, pname, None)
        if val is not None:
            return val

    return 0.0

def show_message(title, message):
    try:
        TaskDialog.Show(title, message)
    except:
        pass


def safe_set_string_param_by_name(element, param_name, value):
    """
    Safely sets a string parameter on an element.
    Returns True if successful, False otherwise.
    """
    try:
        if element is None or not param_name:
            return False

        if value is None:
            return False

        if isinstance(value, str):
            value = value.strip()
            if not value:
                return False
        else:
            value = str(value).strip()
            if not value:
                return False

        param = element.LookupParameter(param_name)
        if not param:
            return False

        if param.IsReadOnly:
            return False

        if param.StorageType == DB.StorageType.String:
            param.Set(value)
            return True

        # fallback in case helper params are not string
        try:
            set_parameter_by_name(element, param_name, value)
            return True
        except:
            return False

    except:
        return False


def safe_copy_param_as_string(source, target, source_param_name, target_param_name):
    """
    Gets a parameter value as string from source and safely copies it to target.
    Skips blanks and missing params.
    """
    try:
        if source is None or target is None:
            return False

        value = get_parameter_value_by_name_AsString(source, source_param_name)

        if value is None:
            return False

        if isinstance(value, str) and not value.strip():
            return False

        return safe_set_string_param_by_name(target, target_param_name, value)

    except:
        return False


# --------------------------------------------------
# Sleeve length storage
# --------------------------------------------------
folder_name = r"c:\Temp"
filepath = os.path.join(folder_name, 'Ribbon_Sleeve.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)

if not os.path.exists(filepath):
    with open(filepath, 'w') as f:
        f.write('6')

with open(filepath, 'r') as f:
    SleeveLength = float(f.read())

annular_filepath = os.path.join(folder_name, 'Ribbon_Duct-Wall-Sleeve.txt')

if not os.path.exists(annular_filepath):
    with open(annular_filepath, 'w') as f:
        f.write('1')

with open(annular_filepath, 'r') as f:
    AnnularSpace = float(f.read()) / 12.0   # stored in inches, convert to feet

# --------------------------------------------------
# Family load options
# --------------------------------------------------
class FamilyLoaderOptionsHandler(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = False
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Project
        overwriteParameterValues.Value = False
        return True


def shared_family_dialog_fallback(sender, args):
    try:
        if isinstance(args, TaskDialogShowingEventArgs):
            msg = (args.Message or "").lower()
            dialog_id = (args.DialogId or "").lower()

            if ("shared" in msg and "already exists" in msg and "project" in msg) \
               or ("shared" in dialog_id and "family" in dialog_id):
                args.OverrideResult(1003)
    except:
        pass


# --------------------------------------------------
# Family manager
# --------------------------------------------------
class SleeveFamilyManager(object):
    def __init__(self, document, family_name, family_path):
        self.doc = document
        self.family_name = family_name
        self.family_path = family_path
        self.family = None
        self.symbol_cache = {}

    def get_family_by_name(self):
        if self.family and self.family.IsValidObject:
            return self.family

        for fam in FilteredElementCollector(self.doc).OfClass(Family):
            try:
                if fam.Name.strip().upper() == self.family_name.strip().upper():
                    self.family = fam
                    return fam
            except:
                pass

        self.family = None
        return None

    def get_symbol_name(self, symbol):
        try:
            if symbol.Name:
                return symbol.Name
        except:
            pass

        try:
            p = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
            if p:
                return p.AsString()
        except:
            pass

        return None

    def load_family_if_missing(self):
        fam = self.get_family_by_name()
        if fam:
            return fam

        if not os.path.exists(self.family_path):
            raise Exception("Family file not found:\n{}".format(self.family_path))

        t = None
        result = False
        loaded_family_ref = clr.Reference[Family]()

        uiapp.DialogBoxShowing += shared_family_dialog_fallback
        try:
            t = Transaction(self.doc, "Load {} Family".format(self.family_name))
            t.Start()

            result = self.doc.LoadFamily(
                self.family_path,
                FamilyLoaderOptionsHandler(),
                loaded_family_ref
            )

            t.Commit()

        except Exception as e:
            if t and t.HasStarted() and not t.HasEnded():
                t.RollBack()
            raise Exception("Family load error for '{}': {}".format(self.family_path, str(e)))

        finally:
            uiapp.DialogBoxShowing -= shared_family_dialog_fallback
            if t:
                t.Dispose()

        # 1) best case: Revit gives back the family directly
        if loaded_family_ref.Value and loaded_family_ref.Value.IsValidObject:
            self.family = loaded_family_ref.Value
            return self.family

        # 2) fallback: search project again
        fam = self.get_family_by_name()
        if fam:
            self.family = fam
            return fam

        raise Exception(
            "LoadFamily failed for '{}'.\n"
            "Expected family name: '{}'\n"
            "LoadFamily result: {}\n"
            "Returned family ref: {}".format(
                self.family_path,
                self.family_name,
                result,
                loaded_family_ref.Value.Name if loaded_family_ref.Value else "<None>"
            )
        )

    def build_symbol_cache(self):
        self.symbol_cache = {}

        fam = self.get_family_by_name()
        if not fam:
            return

        for symbol_id in fam.GetFamilySymbolIds():
            sym = self.doc.GetElement(symbol_id)
            if sym:
                type_name = self.get_symbol_name(sym)
                if type_name:
                    self.symbol_cache[type_name.strip().upper()] = sym

    def get_available_type_names(self):
        if not self.symbol_cache:
            self.build_symbol_cache()
        return sorted(self.symbol_cache.keys())

    def get_symbol_by_type_name(self, type_name=None):
        if not self.symbol_cache:
            self.build_symbol_cache()

        if type_name:
            sym = self.symbol_cache.get(type_name.strip().upper())
            if sym:
                return sym

        for k in sorted(self.symbol_cache.keys()):
            return self.symbol_cache[k]

        return None

    def activate_symbol_if_needed(self, symbol):
        if not symbol:
            return False

        if symbol.IsActive:
            return True

        t = None
        try:
            t = Transaction(self.doc, "Activate {} Symbol".format(self.family_name))
            t.Start()
            symbol.Activate()
            self.doc.Regenerate()
            t.Commit()
            return True
        except Exception as e:
            if t and t.HasStarted() and not t.HasEnded():
                t.RollBack()
            show_message("Error", "Family symbol activation error:\n{}".format(str(e)))
            return False
        finally:
            if t:
                t.Dispose()

    def get_ready_symbol(self, preferred_type_name=None):
        fam = self.load_family_if_missing()
        if not fam:
            show_message("Error", "Family '{}' was not found in project after load.".format(self.family_name))
            return None

        self.build_symbol_cache()

        sym = self.get_symbol_by_type_name(preferred_type_name)
        if not sym:
            available = self.get_available_type_names()
            show_message(
                "Error",
                "No usable type found in family '{}'.\nAvailable types: {}".format(
                    self.family_name,
                    ", ".join(available) if available else "<none>"
                )
            )
            return None

        if not self.activate_symbol_if_needed(sym):
            return None

        return sym


round_family_manager = SleeveFamilyManager(doc, 'Round Wall Sleeve', round_family_path)
rect_family_manager = SleeveFamilyManager(doc, 'Rectangular Wall Sleeve', rect_family_path)

round_symbol = round_family_manager.get_ready_symbol()
rect_symbol = rect_family_manager.get_ready_symbol()

if not round_symbol:
    raise Exception("Round Wall Sleeve symbol not found or could not be activated.")

if not rect_symbol:
    raise Exception("Rectangular Wall Sleeve symbol not found or could not be activated.")


# --------------------------------------------------
# Selection / geometry
# --------------------------------------------------

class FabricationDuctSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        try:
            if elem is None:
                return False

            if not isinstance(elem, DB.FabricationPart):
                return False

            if elem.Category is None:
                return False

            fab_duct_cat_id = DB.ElementId(DB.BuiltInCategory.OST_FabricationDuctwork).IntegerValue
            return elem.Category.Id.IntegerValue == fab_duct_cat_id

        except:
            return False

    def AllowReference(self, reference, point):
        return False

def select_fabrication_duct():
    selection_filter = FabricationDuctSelectionFilter()

    ref = uidoc.Selection.PickObject(
        ObjectType.Element,
        selection_filter,
        "Select fabrication ductwork only"
    )

    return doc.GetElement(ref.ElementId)


def pick_point():
    return uidoc.Selection.PickPoint("Pick a point along the centerline of the duct")


def get_duct_centerline(duct):
    duct_location = duct.Location
    if isinstance(duct_location, LocationCurve):
        return duct_location.Curve
    raise Exception("The selected element does not have a valid centerline.")


def project_point_on_curve(point, curve):
    result = curve.Project(point)
    if not result:
        raise Exception("Could not project point onto duct centerline.")
    return result.XYZPoint


def create_family_instance(insertion_point, symbol):
    try:
        if level:
            return doc.Create.NewFamilyInstance(
                insertion_point,
                symbol,
                level,
                DB.Structure.StructuralType.NonStructural
            )
    except:
        pass

    return doc.Create.NewFamilyInstance(
        insertion_point,
        symbol,
        DB.Structure.StructuralType.NonStructural
    )


def get_connectors(element):
    try:
        return list(element.ConnectorManager.Connectors)
    except:
        raise Exception("Could not read connectors from selected duct.")


def get_shape_and_size_from_connector(connector, duct):
    shape = connector.Shape
    insulation_thickness = get_insulation_thickness(duct)

    if shape == DB.ConnectorProfileType.Round:
        duct_diameter = connector.Radius * 2.0
        sleeve_diameter = duct_diameter + (2.0 * insulation_thickness) + AnnularSpace

        return "ROUND", {
            "Diameter": sleeve_diameter
        }

    elif shape == DB.ConnectorProfileType.Rectangular:
        duct_width = connector.Width
        duct_height = connector.Height

        sleeve_width = duct_width + (2.0 * insulation_thickness) + AnnularSpace
        sleeve_height = duct_height + (2.0 * insulation_thickness) + AnnularSpace

        return "RECTANGULAR", {
            "Width": sleeve_width,
            "Height": sleeve_height
        }

    raise Exception("Only round and rectangular fabrication ductwork are supported.")


def copy_fab_params(source, target):
    safe_copy_param_as_string(source, target, 'Overall Size', 'FP_Product Entry')
    safe_copy_param_as_string(source, target, 'Fabrication Service Name', 'FP_Service Name')
    safe_copy_param_as_string(source, target, 'Fabrication Service Abbreviation', 'FP_Service Abbreviation')


def rotate_to_duct(instance, connectors, insertion_point, picked_point):
    if len(connectors) < 2:
        return

    conn1, conn2 = connectors[0], connectors[1]
    nearest_conn = min([conn1, conn2], key=lambda c: picked_point.DistanceTo(c.Origin))
    other_conn = conn2 if nearest_conn == conn1 else conn1

    vec = other_conn.Origin - nearest_conn.Origin
    angle = atan2(vec.Y, vec.X)

    axis = DB.Line.CreateBound(
        insertion_point,
        DB.XYZ(insertion_point.X, insertion_point.Y, insertion_point.Z + 1.0)
    )
    DB.ElementTransformUtils.RotateElement(doc, instance.Id, axis, angle)


def place_and_modify_family(duct):
    centerline_curve = get_duct_centerline(duct)
    picked_point = pick_point()
    projected_point = project_point_on_curve(picked_point, centerline_curve)
    insertion_point = DB.XYZ(picked_point.X, picked_point.Y, projected_point.Z)

    connectors = get_connectors(duct)
    if len(connectors) < 2:
        raise Exception("Selected duct does not have enough connectors.")

    nearest_conn = min(connectors, key=lambda c: picked_point.DistanceTo(c.Origin))
    shape_name, size_data = get_shape_and_size_from_connector(nearest_conn, duct)

    if shape_name == "ROUND":
        symbol = round_symbol
    elif shape_name == "RECTANGULAR":
        symbol = rect_symbol
    else:
        raise Exception("Unsupported duct shape.")

    new_family_instance = create_family_instance(insertion_point, symbol)
    if not new_family_instance:
        raise Exception("Failed to create family instance.")

    if shape_name == "ROUND":
        set_parameter_by_name(new_family_instance, 'Diameter', size_data['Diameter'])

    elif shape_name == "RECTANGULAR":
        set_parameter_by_name(new_family_instance, 'Width', size_data['Width'])
        set_parameter_by_name(new_family_instance, 'Height', size_data['Height'])

    set_parameter_by_name(new_family_instance, 'Length', SleeveLength)

    rotate_to_duct(new_family_instance, connectors, insertion_point, picked_point)
    copy_fab_params(duct, new_family_instance)

    schedule_level_param = new_family_instance.LookupParameter("Schedule Level")
    if schedule_level_param and not schedule_level_param.IsReadOnly and level:
        schedule_level_param.Set(level.Id)


# --------------------------------------------------
# Main execution loop
# --------------------------------------------------
while True:
    t = None
    try:
        duct = select_fabrication_duct()

        t = Transaction(doc, 'Place Wall Sleeve Family')
        t.Start()
        place_and_modify_family(duct)
        t.Commit()

    except OperationCanceledException:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        break

    except Exception as e:
        if t and t.HasStarted() and not t.HasEnded():
            t.RollBack()
        show_message("Error", str(e))
        break

    finally:
        if t:
            t.Dispose()