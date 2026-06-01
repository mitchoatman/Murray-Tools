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
    Transaction,
    Wall
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
    active_view_level = active_view.GenLevel
except:
    active_view_level = None


folder_name = r"c:\Temp"
annular_filepath = os.path.join(folder_name, 'Ribbon_Duct-Wall-Sleeve.txt')

if not os.path.exists(folder_name):
    os.makedirs(folder_name)

if not os.path.exists(annular_filepath):
    with open(annular_filepath, 'w') as f:
        f.write('1')

with open(annular_filepath, 'r') as f:
    AnnularSpace = float(f.read()) / 12.0

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
    try:
        if element is None or not param_name or value is None:
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
        if not param or param.IsReadOnly:
            return False

        if param.StorageType == DB.StorageType.String:
            param.Set(value)
            return True

        return False
    except:
        return False


def safe_copy_param_as_string(source, target, source_param_name, target_param_name):
    try:
        value = get_parameter_value_by_name_AsString(source, source_param_name)
        if value is None:
            return False
        if isinstance(value, str) and not value.strip():
            return False
        return safe_set_string_param_by_name(target, target_param_name, value)
    except:
        return False


def copy_fab_params(source, target):
    safe_copy_param_as_string(source, target, 'Overall Size', 'FP_Product Entry')
    safe_copy_param_as_string(source, target, 'Fabrication Service Name', 'FP_Service Name')
    safe_copy_param_as_string(source, target, 'Fabrication Service Abbreviation', 'FP_Service Abbreviation')


def get_best_level_id(element):
    try:
        if element.LevelId and element.LevelId != DB.ElementId.InvalidElementId:
            return element.LevelId
    except:
        pass

    try:
        if active_view_level:
            return active_view_level.Id
    except:
        pass

    return None


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
            show_message("Error", "Family file not found:\n{}".format(self.family_path))
            return None

        t = None
        uiapp.DialogBoxShowing += shared_family_dialog_fallback
        try:
            t = Transaction(self.doc, "Load {} Family".format(self.family_name))
            t.Start()

            loaded_family_ref = clr.Reference[Family]()
            result = self.doc.LoadFamily(
                self.family_path,
                FamilyLoaderOptionsHandler(),
                loaded_family_ref
            )

            t.Commit()

            fam = self.get_family_by_name()
            if fam:
                self.family = fam
                return fam

            if result and loaded_family_ref.Value:
                self.family = loaded_family_ref.Value
                return self.family

            return None

        except Exception as e:
            if t and t.HasStarted() and not t.HasEnded():
                t.RollBack()
            show_message("Error", "Family load error:\n{}".format(str(e)))
            return None

        finally:
            uiapp.DialogBoxShowing -= shared_family_dialog_fallback
            if t:
                t.Dispose()

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
            return None

        self.build_symbol_cache()
        sym = self.get_symbol_by_type_name(preferred_type_name)

        if not sym:
            show_message("Error", "No usable type found in family '{}'.".format(self.family_name))
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
# Selection filters
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


class LinkedWallSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, DB.RevitLinkInstance)

    def AllowReference(self, reference, position):
        try:
            link_instance = doc.GetElement(reference.ElementId)
            if not isinstance(link_instance, DB.RevitLinkInstance):
                return False

            linked_doc = link_instance.GetLinkDocument()
            if linked_doc is None:
                return False

            linked_elem = linked_doc.GetElement(reference.LinkedElementId)
            return isinstance(linked_elem, DB.Wall)
        except:
            return False


# --------------------------------------------------
# Selection
# --------------------------------------------------
def select_fabrication_ducts():
    refs = uidoc.Selection.PickObjects(
        ObjectType.Element,
        FabricationDuctSelectionFilter(),
        "Select one or more fabrication ducts, then press Finish"
    )

    ducts = [doc.GetElement(r.ElementId) for r in refs]
    if not ducts:
        raise Exception("No fabrication ducts selected.")

    return ducts


def select_linked_wall():
    return uidoc.Selection.PickObject(
        ObjectType.LinkedElement,
        LinkedWallSelectionFilter(),
        "Select a wall in a Revit link"
    )


# --------------------------------------------------
# Geometry
# --------------------------------------------------
def get_duct_centerline(duct):
    loc = duct.Location
    if isinstance(loc, LocationCurve):
        return loc.Curve
    raise Exception("Selected duct does not have a valid centerline.")


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


def get_linked_wall_thickness_and_curve(link_ref):
    link_instance = doc.GetElement(link_ref.ElementId)
    if not link_instance:
        raise Exception("Could not get Revit link instance.")

    linked_doc = link_instance.GetLinkDocument()
    if linked_doc is None:
        raise Exception("Could not access linked document.")

    wall = linked_doc.GetElement(link_ref.LinkedElementId)
    if not isinstance(wall, Wall):
        raise Exception("Selected linked element is not a wall.")

    wall_type = wall.WallType
    thickness = wall_type.get_Parameter(DB.BuiltInParameter.WALL_ATTR_WIDTH_PARAM).AsDouble()

    loc = wall.Location
    if not isinstance(loc, LocationCurve):
        raise Exception("Selected wall does not have a valid location curve.")

    wall_curve = loc.Curve

    try:
        transform = link_instance.GetTotalTransform()
    except:
        transform = link_instance.GetTransform()

    transformed_wall_curve = wall_curve.CreateTransformed(transform)
    return thickness, transformed_wall_curve


def project_wall_curve_to_duct_plane(wall_curve, duct_curve):
    duct_start = duct_curve.GetEndPoint(0)
    duct_end = duct_curve.GetEndPoint(1)
    duct_direction = (duct_end - duct_start).Normalize()

    plane_normal = DB.XYZ(0, 0, 1)
    if abs(duct_direction.Z) > 0.99:
        plane_normal = DB.XYZ(1, 0, 0)

    plane = DB.Plane.CreateByNormalAndOrigin(plane_normal, duct_start)

    wall_start = wall_curve.GetEndPoint(0)
    wall_end = wall_curve.GetEndPoint(1)

    uv_start, dist_start = plane.Project(wall_start)
    uv_end, dist_end = plane.Project(wall_end)

    origin = plane.Origin
    x_axis = plane.XVec
    y_axis = plane.YVec

    projected_start = origin + x_axis * uv_start.U + y_axis * uv_start.V
    projected_end = origin + x_axis * uv_end.U + y_axis * uv_end.V

    return DB.Line.CreateBound(projected_start, projected_end)


def get_intersection_point(duct_curve, wall_curve):
    result_array = clr.Reference[DB.IntersectionResultArray]()
    result = duct_curve.Intersect(wall_curve, result_array)

    if result != DB.SetComparisonResult.Overlap:
        raise Exception("Duct does not intersect the selected linked wall.")

    if result_array.Value and result_array.Value.Size > 0:
        return result_array.Value[0].XYZPoint

    raise Exception("Failed to retrieve intersection point.")


def create_family_instance(insertion_point, symbol, level_id=None):
    try:
        if level_id and level_id != DB.ElementId.InvalidElementId:
            lvl = doc.GetElement(level_id)
            if lvl:
                return doc.Create.NewFamilyInstance(
                    insertion_point,
                    symbol,
                    lvl,
                    DB.Structure.StructuralType.NonStructural
                )
    except:
        pass

    return doc.Create.NewFamilyInstance(
        insertion_point,
        symbol,
        DB.Structure.StructuralType.NonStructural
    )


def rotate_to_duct(instance, connectors, insertion_point, reference_point):
    if len(connectors) < 2:
        return

    conn1, conn2 = connectors[0], connectors[1]
    nearest_conn = min([conn1, conn2], key=lambda c: reference_point.DistanceTo(c.Origin))
    other_conn = conn2 if nearest_conn == conn1 else conn1

    vec = other_conn.Origin - nearest_conn.Origin
    angle = atan2(vec.Y, vec.X)

    axis = DB.Line.CreateBound(
        insertion_point,
        DB.XYZ(insertion_point.X, insertion_point.Y, insertion_point.Z + 1.0)
    )

    DB.ElementTransformUtils.RotateElement(doc, instance.Id, axis, angle)


# --------------------------------------------------
# Placement
# --------------------------------------------------
def place_linked_wall_sleeve(duct, wall_ref):
    duct_curve = get_duct_centerline(duct)
    wall_thickness, wall_curve = get_linked_wall_thickness_and_curve(wall_ref)

    projected_wall_curve = project_wall_curve_to_duct_plane(wall_curve, duct_curve)
    intersection_point = get_intersection_point(duct_curve, projected_wall_curve)

    connectors = get_connectors(duct)
    if len(connectors) < 2:
        raise Exception("Selected duct does not have enough connectors.")

    nearest_conn = min(connectors, key=lambda c: intersection_point.DistanceTo(c.Origin))
    other_conn = [c for c in connectors if c.Id != nearest_conn.Id][0]
    duct_direction = (other_conn.Origin - nearest_conn.Origin).Normalize()

    # keeps same behavior as your linked-pipe version:
    # inserts sleeve half wall thickness back from centerline intersection
    insertion_point = intersection_point - duct_direction * (wall_thickness / 2.0)

    shape_name, size_data = get_shape_and_size_from_connector(nearest_conn, duct)

    if shape_name == "ROUND":
        symbol = round_symbol
    elif shape_name == "RECTANGULAR":
        symbol = rect_symbol
    else:
        raise Exception("Unsupported duct shape.")

    level_id = get_best_level_id(duct)
    new_family_instance = create_family_instance(insertion_point, symbol, level_id)

    if not new_family_instance:
        raise Exception("Failed to create family instance.")

    if shape_name == "ROUND":
        set_parameter_by_name(new_family_instance, 'Diameter', size_data['Diameter'])

    elif shape_name == "RECTANGULAR":
        set_parameter_by_name(new_family_instance, 'Width', size_data['Width'])
        set_parameter_by_name(new_family_instance, 'Height', size_data['Height'])

    set_parameter_by_name(new_family_instance, 'Length', wall_thickness)

    rotate_to_duct(new_family_instance, connectors, insertion_point, intersection_point)
    copy_fab_params(duct, new_family_instance)

    schedule_level_param = new_family_instance.LookupParameter("Schedule Level")
    if schedule_level_param and not schedule_level_param.IsReadOnly and level_id:
        schedule_level_param.Set(level_id)


# --------------------------------------------------
# Main
# --------------------------------------------------
failed = []
t = None

try:
    ducts = select_fabrication_ducts()
    wall_ref = select_linked_wall()

    t = Transaction(doc, 'Place Linked Wall Sleeve Family')
    t.Start()

    for duct in ducts:
        try:
            place_linked_wall_sleeve(duct, wall_ref)
        except Exception as duct_error:
            try:
                failed.append("Duct {}: {}".format(duct.Id.IntegerValue, str(duct_error)))
            except:
                failed.append(str(duct_error))

    if failed and len(failed) == len(ducts):
        t.RollBack()
        show_message("Wall Sleeve Placement", "No sleeves were placed.\n\n{}".format("\n".join(failed[:20])))
    else:
        t.Commit()
        if failed:
            show_message(
                "Wall Sleeve Placement",
                "Placement completed with some issues.\n\n{}".format("\n".join(failed[:20]))
            )

except OperationCanceledException:
    if t and t.HasStarted() and not t.HasEnded():
        t.RollBack()

except Exception as e:
    if t and t.HasStarted() and not t.HasEnded():
        t.RollBack()
    show_message("Error", "Error during execution:\n{}".format(str(e)))

finally:
    if t:
        t.Dispose()